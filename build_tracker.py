"""
State Power Transition Tracker — full rebuild pipeline.

Run this whenever the tracker artifact needs refreshing. It:
  1. Downloads the latest EIA "Net Generation by State by Type of Producer
     by Energy Source" annual workbook (no API key required) — final data,
     currently through 2024.
  2. Downloads EIA's companion monthly workbook (also no API key) and, for
     any full calendar year not yet covered by the annual "Final" release
     (currently 2025), rolls up all twelve months into a real annual actual
     per state. This is EIA's own preliminary monthly data, not a modeled
     estimate — it gets merged in as an actual year, flagged 'prelim' so the
     footer/tooltips can note it isn't the finalized EIA-923 figure yet.
  3. Recomputes each state's fossil / renewable / nuclear generation share
     for every available year.
  4. Merges in state_goals.json (hand-curated RPS/CES/100%-mandate status —
     edit that file directly if a state's law changes; this script does not
     rewrite it).
  5. Forecasts each state's renewable share out to 2032 by fitting a logistic
     (S-curve) to that state's own actual history — the ceiling, steepness,
     and inflection year all come from the data. A state's legal mandate
     (from state_goals.json) never bends the fit; if a state's actual
     trajectory implies it won't reach its mandate, the forecast says so.
     States without enough curvature signal for a stable logistic fit fall
     back to a bounded linear projection and get flagged low-confidence.
  6. Renders the self-contained artifact.html in this same folder.

After running this script, read artifact.html back and pass it to
mcp__cowork__update_artifact with id "state-power-transition-tracker".

Usage:  python3 build_tracker.py
"""
import json
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).parent
XLS_URL = "https://www.eia.gov/electricity/data/state/annual_generation_state.xls"
XLS_MONTHLY_URL = "https://www.eia.gov/electricity/data/state/generation_monthly.xlsx"
FORECAST_THROUGH_YEAR = 2032

STATE_NAMES = {
 'AL':'Alabama','AK':'Alaska','AZ':'Arizona','AR':'Arkansas','CA':'California','CO':'Colorado',
 'CT':'Connecticut','DE':'Delaware','DC':'District of Columbia','FL':'Florida','GA':'Georgia',
 'HI':'Hawaii','ID':'Idaho','IL':'Illinois','IN':'Indiana','IA':'Iowa','KS':'Kansas','KY':'Kentucky',
 'LA':'Louisiana','ME':'Maine','MD':'Maryland','MA':'Massachusetts','MI':'Michigan','MN':'Minnesota',
 'MS':'Mississippi','MO':'Missouri','MT':'Montana','NE':'Nebraska','NV':'Nevada','NH':'New Hampshire',
 'NJ':'New Jersey','NM':'New Mexico','NY':'New York','NC':'North Carolina','ND':'North Dakota',
 'OH':'Ohio','OK':'Oklahoma','OR':'Oregon','PA':'Pennsylvania','RI':'Rhode Island','SC':'South Carolina',
 'SD':'South Dakota','TN':'Tennessee','TX':'Texas','UT':'Utah','VT':'Vermont','VA':'Virginia',
 'WA':'Washington','WV':'West Virginia','WI':'Wisconsin','WY':'Wyoming'
}
FOSSIL = {'Coal','Natural Gas','Petroleum','Other Gases'}
RENEWABLE = {'Hydroelectric Conventional','Wind','Solar Thermal and Photovoltaic','Geothermal',
             'Wood and Wood Derived Fuels','Other Biomass'}
NUCLEAR = {'Nuclear'}


def fetch_xls():
    dest = HERE / 'gen_state.xlsx'
    subprocess.run(['curl', '-sL', '-o', str(dest), XLS_URL, '-A', 'Mozilla/5.0'], check=True)
    return dest


def fetch_monthly_xls():
    dest = HERE / 'gen_state_monthly.xlsx'
    subprocess.run(['curl', '-sL', '-o', str(dest), XLS_MONTHLY_URL, '-A', 'Mozilla/5.0'], check=True)
    return dest


def process_monthly_rollup(xlsx_path, year):
    """Roll up EIA's monthly workbook into a real annual actual for `year`.

    Returns a dict {state_code: {'t','f','r','n'}} — same shape as one year's
    entry from process_generation() — or None if the sheet for `year` isn't
    present, or all twelve months aren't yet posted (a partial year must
    never be merged in as if it were a full-year actual).
    """
    import pandas as pd

    sheet_candidates = [f'{year}_Preliminary', f'{year}_Final']
    xl = pd.ExcelFile(xlsx_path)
    sheet = next((s for s in sheet_candidates if s in xl.sheet_names), None)
    if sheet is None:
        return None

    df = pd.read_excel(xlsx_path, sheet_name=sheet, header=4)
    df.columns = ['year', 'month', 'state', 'producer', 'source', 'mwh']
    months_present = set(int(m) for m in df['month'].unique())
    if months_present != set(range(1, 13)):
        print(f"  monthly rollup for {year}: only {len(months_present)}/12 months posted — skipping "
              f"(partial year, not a valid annual actual)")
        return None

    df['state'] = df['state'].astype(str).str.strip()
    d = df[(df.producer == 'Total Electric Power Industry') & (df.state.isin(STATE_NAMES.keys()))].copy()

    out = {}
    for state, g in d.groupby('state'):
        total = g[g.source == 'Total']['mwh'].sum()
        if not total or total <= 0:
            continue
        fossil = g[g.source.isin(FOSSIL)]['mwh'].sum()
        renew = g[g.source.isin(RENEWABLE)]['mwh'].sum()
        nuke = g[g.source.isin(NUCLEAR)]['mwh'].sum()
        out[state] = {
            't': round(total),
            'f': round(fossil / total * 1000) / 10,
            'r': round(renew / total * 1000) / 10,
            'n': round(nuke / total * 1000) / 10,
        }
    return out


def process_generation(xls_path):
    import pandas as pd
    df = pd.read_excel(xls_path, sheet_name=0, header=1)
    df.columns = ['year', 'state', 'producer', 'source', 'mwh']
    df['state'] = df['state'].astype(str).str.strip()
    d = df[(df.producer == 'Total Electric Power Industry') & (df.state.isin(STATE_NAMES.keys()))].copy()

    out = {}
    for state, g in d.groupby('state'):
        out[state] = {'name': STATE_NAMES[state], 'years': {}}
        for year, gy in g.groupby('year'):
            total = gy[gy.source == 'Total']['mwh'].sum()
            if not total or total <= 0:
                continue
            fossil = gy[gy.source.isin(FOSSIL)]['mwh'].sum()
            renew = gy[gy.source.isin(RENEWABLE)]['mwh'].sum()
            nuke = gy[gy.source.isin(NUCLEAR)]['mwh'].sum()
            out[state]['years'][str(int(year))] = {
                't': round(total),
                'f': round(fossil / total * 1000) / 10,
                'r': round(renew / total * 1000) / 10,
                'n': round(nuke / total * 1000) / 10,
            }
    return out


def _logistic(t, L, k, t0):
    import numpy as np
    return L / (1 + np.exp(-k * (t - t0)))


def fit_logistic_share(years, values):
    """Fit a logistic S-curve to a state's actual renewable-share history.

    `years`/`values` are the full actual series (no forecast years mixed in).
    The ceiling (L), steepness (k), and inflection year (t0) are all fit from
    this state's own data — nothing here reads state_goals.json or any
    mandate. Returns {'converged': False} if there's too little history or
    the fit doesn't converge; caller falls back to a linear projection.
    """
    import numpy as np
    from scipy.optimize import curve_fit

    if len(years) < 6:
        return {'converged': False}

    yrs = np.array(years, dtype=float)
    vals = np.array(values, dtype=float)
    cur_val = vals[-1]
    t0_guess = yrs[len(yrs) // 2]
    L_guess = min(100.0, max(cur_val + 15.0, 40.0))
    p0 = [L_guess, 0.15, t0_guess]
    # L can't be forced below the state's current actual share, and can't
    # exceed 100%. k and t0 are bounded generously — they describe the shape
    # of this state's own curve, not a target.
    bounds = ([max(cur_val, 0.5), 0.01, yrs.min() - 40], [100.0, 1.5, yrs.max() + 60])

    try:
        popt, _ = curve_fit(_logistic, yrs, vals, p0=p0, bounds=bounds, maxfev=8000)
    except Exception:
        return {'converged': False}

    L, k, t0 = (float(x) for x in popt)
    pred = _logistic(yrs, L, k, t0)
    ss_res = float(np.sum((vals - pred) ** 2))
    ss_tot = float(np.sum((vals - vals.mean()) ** 2))
    r2 = 1.0 - ss_res / ss_tot if ss_tot > 0 else 0.0
    return {'converged': True, 'L': L, 'k': k, 't0': t0, 'r2': r2}


MIN_LOGISTIC_R2 = 0.75


def linear_fallback(years, values, target_years):
    """Bounded linear projection for states with no stable S-curve signal.

    Uses at most the last 10 actual years to estimate a slope (an old,
    no-longer-representative trend shouldn't dominate), but anchors the
    projection at the last actual value rather than the raw OLS line. A
    noisy series (e.g. a hydro-heavy state's year-to-year swings) can sit
    well off its own best-fit line at the most recent point; anchoring at
    the fitted line instead of the last actual would produce a visible
    jump at the actual/forecast boundary that isn't a real trend, just an
    artifact of where the last data point happened to fall. Clipped to
    [0, 100] — a linear projection can wander outside a valid share range
    if extrapolated far enough.
    """
    import numpy as np
    n = min(10, len(years))
    yrs = np.array(years[-n:], dtype=float)
    vals = np.array(values[-n:], dtype=float)
    slope = float(np.polyfit(yrs, vals, 1)[0]) if n >= 2 else 0.0
    last_val = values[-1]
    last_year = years[-1]
    out = {}
    for ty in target_years:
        v = last_val + slope * (ty - last_year)
        out[ty] = max(0.0, min(100.0, float(v)))
    return out


def forecast_nuclear_share(years, values, target_years, damping=0.3):
    """Nuclear isn't undergoing a diffusion transition in most states, so an
    S-curve is the wrong shape for it. Extrapolate the last-10-year trend at
    30% strength (damped, not flat — allows for a state's known retirement or
    new-build schedule to show up gradually without overreacting to a couple
    of noisy years), clipped to a valid share range.
    """
    import numpy as np
    n = min(10, len(years))
    yrs = np.array(years[-n:], dtype=float)
    vals = np.array(values[-n:], dtype=float)
    slope = float(np.polyfit(yrs, vals, 1)[0]) if n >= 3 else 0.0
    last_val = values[-1]
    last_year = years[-1]
    out = {}
    for ty in target_years:
        v = last_val + damping * slope * (ty - last_year)
        out[ty] = max(0.0, min(100.0, v))
    return out


def forecast_total_generation(years, totals, target_years, max_annual_growth=0.05):
    """Project each state's total MWh using its own historical CAGR over the
    last 10 actual years, capped at +/-5%/year so a short noisy run of years
    can't compound into an implausible total by 2032.
    """
    n = min(10, len(years))
    y0, y1 = years[-n], years[-1]
    t0v, t1v = totals[-n], totals[-1]
    span = max(1, y1 - y0)
    cagr = (t1v / t0v) ** (1.0 / span) - 1.0 if t0v > 0 else 0.0
    cagr = max(-max_annual_growth, min(max_annual_growth, cagr))
    out = {}
    for ty in target_years:
        out[ty] = t1v * ((1.0 + cagr) ** (ty - y1))
    return out


def build_state_forecast(years_dict, target_years):
    """Build the 2026-2032 (or whatever target_years is) forecast for one
    state. Returns (forecast_series, fit_meta):
      forecast_series: {year_str: {'t','f','r','n','p':True}}
      fit_meta: {'method','lowConfidence', + ceiling/k/t0/r2 if logistic}

    Renewable share drives the fit; nuclear gets its own damped trend;
    fossil is whatever's left over (100 - renewable - nuclear, clipped).
    state_goals.json is never consulted here — see module docstring.
    """
    yrs_sorted = sorted(years_dict.keys(), key=int)
    yrs_int = [int(y) for y in yrs_sorted]
    r_vals = [years_dict[y]['r'] for y in yrs_sorted]
    n_vals = [years_dict[y]['n'] for y in yrs_sorted]
    t_vals = [years_dict[y]['t'] for y in yrs_sorted]

    fit = fit_logistic_share(yrs_int, r_vals)
    if fit.get('converged') and fit.get('r2', 0.0) >= MIN_LOGISTIC_R2:
        r_forecast = {ty: max(0.0, min(100.0, _logistic(ty, fit['L'], fit['k'], fit['t0']))) for ty in target_years}
        fit_meta = {
            'method': 'logistic', 'lowConfidence': False,
            'ceiling': round(fit['L'], 1), 'k': round(fit['k'], 3),
            't0': round(fit['t0'], 1), 'r2': round(fit['r2'], 3),
        }
    else:
        r_forecast = linear_fallback(yrs_int, r_vals, target_years)
        fit_meta = {'method': 'linear_fallback', 'lowConfidence': True}

    n_forecast = forecast_nuclear_share(yrs_int, n_vals, target_years)
    t_forecast = forecast_total_generation(yrs_int, t_vals, target_years)

    forecast_series = {}
    for ty in target_years:
        r = r_forecast[ty]
        n = n_forecast[ty]
        if r + n > 100.0:
            r = max(0.0, 100.0 - n)  # renewable yields to nuclear, not the other way around
        f = round(100.0 - r - n, 1)
        forecast_series[str(ty)] = {
            't': round(t_forecast[ty]),
            'f': f,
            'r': round(r, 1),
            'n': round(n, 1),
            'p': True,
        }
    return forecast_series, fit_meta


def build_site_data(gen, goals, last_annual_final_year, prelim_years=(), forecast_through=FORECAST_THROUGH_YEAR):
    prelim_years = set(prelim_years)
    combined = {}
    for code, g in gen.items():
        if code not in goals:
            continue
        years = g['years']
        yrs_sorted = sorted(years.keys(), key=int)
        if not yrs_sorted:
            continue
        last_yr = yrs_sorted[-1]
        target_prior = str(int(last_yr) - 10)
        prior_yr = target_prior if target_prior in years else yrs_sorted[max(0, len(yrs_sorted) - 11)]
        cur = years[last_yr]
        prior = years[prior_yr]
        combined[code] = {
            'name': g['name'],
            'goal': goals[code],
            'series': {y: years[y] for y in yrs_sorted if int(y) >= 2001},
            'cur': {'year': int(last_yr), 'f': cur['f'], 'r': cur['r'], 'n': cur['n'], 't': cur['t'],
                    'prelim': last_yr in prelim_years},
            'd10': {'from_year': int(prior_yr), 'df': round(cur['f'] - prior['f'], 1), 'dr': round(cur['r'] - prior['r'], 1)}
        }

    last_actual_data_year = max(int(y) for s in combined.values() for y in s['series'].keys())

    # Forecast 2026-2032 (or whatever forecast_through is) per state. Runs
    # after 'cur'/'d10' are locked in above so the headline stats on each
    # card stay pinned to the latest actual year, never the forecast.
    target_years = list(range(last_actual_data_year + 1, forecast_through + 1))
    if target_years:
        for code, s in combined.items():
            forecast_series, fit_meta = build_state_forecast(gen[code]['years'], target_years)
            s['series'].update(forecast_series)
            s['forecast'] = fit_meta

    years_all = sorted({y for s in combined.values() for y in s['series'].keys()}, key=int)
    nat = {}
    for y in years_all:
        tot = fos = ren = nuc = 0
        for s in combined.values():
            if y in s['series']:
                row = s['series'][y]
                t = row['t']
                tot += t
                fos += t * row['f'] / 100
                ren += t * row['r'] / 100
                nuc += t * row['n'] / 100
        if tot:
            nat[y] = {'f': round(fos / tot * 1000) / 10, 'r': round(ren / tot * 1000) / 10, 'n': round(nuc / tot * 1000) / 10}

    # EIA has historically released the annual "Final" workbook in September,
    # covering the prior calendar year, with the following year's edition
    # around October. That cadence only describes last_annual_final_year —
    # any later year merged in from the monthly rollup is EIA's own
    # preliminary monthly data, not yet the finalized annual release.
    released = f"September {last_annual_final_year + 1}"
    next_release = f"October {last_annual_final_year + 2}"
    return {
        'meta': {
            'lastDataYear': last_actual_data_year,
            'lastAnnualFinalYear': last_annual_final_year,
            'prelimYears': sorted(prelim_years, key=int),
            'forecastThrough': forecast_through if target_years else None,
            'forecastMethod': 'Logistic (S-curve) fit per state on actual renewable-share history; '
                               'ceiling, steepness, and inflection year are all data-derived, not tied to '
                               'any state\'s legal RPS/CES mandate. Nuclear share uses a damped historical '
                               'trend; fossil share is the remainder. States without enough curvature signal '
                               'for a stable fit fall back to a bounded linear projection and are flagged '
                               'low-confidence.',
            'released': released,
            'nextRelease': next_release,
            'source': 'EIA, Net Generation by State by Type of Producer by Energy Source (1990–present); '
                      'years after the annual Final release are rolled up from EIA\'s monthly Electric Power '
                      'Monthly data and are preliminary until EIA\'s next annual Final release.',
        },
        'national': nat,
        'states': combined,
    }


HTML_TEMPLATE_PATH = HERE / 'template.html'


def render_html(site_data):
    template = HTML_TEMPLATE_PATH.read_text()
    data_json = json.dumps(site_data, separators=(',', ':'))
    html = template.replace('__DATA_JSON__', data_json)
    out_path = HERE / 'artifact.html'
    out_path.write_text(html)
    return out_path


def main():
    goals = json.loads((HERE / 'state_goals.json').read_text())

    xls_path = fetch_xls()
    gen = process_generation(xls_path)
    last_annual_final_year = max(int(y) for s in gen.values() for y in s['years'])

    # Extend with real actuals for every full calendar year after the annual
    # "Final" release, using EIA's own monthly workbook — no forecasting yet.
    monthly_path = fetch_monthly_xls()
    prelim_years = []
    candidate_year = last_annual_final_year + 1
    while True:
        rollup = process_monthly_rollup(monthly_path, candidate_year)
        if rollup is None:
            break
        for code, row in rollup.items():
            if code in gen:
                gen[code]['years'][str(candidate_year)] = row
        prelim_years.append(str(candidate_year))
        print(f"  merged {candidate_year} actuals from EIA monthly rollup ({len(rollup)} states) — preliminary")
        candidate_year += 1

    site_data = build_site_data(gen, goals, last_annual_final_year, prelim_years)
    out_path = render_html(site_data)
    print(f"Built {out_path} — data through {site_data['meta']['lastDataYear']} "
          f"(annual Final through {last_annual_final_year}; preliminary: {', '.join(prelim_years) or 'none'}), "
          f"{len(site_data['states'])} states.")


if __name__ == '__main__':
    main()
