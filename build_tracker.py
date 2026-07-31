"""
State Power Transition Tracker — full rebuild pipeline.

Run this whenever the tracker artifact needs refreshing. It:
  1. Downloads the latest EIA "Net Generation by State by Type of Producer
     by Energy Source" workbook (no API key required).
  2. Recomputes each state's fossil / renewable / nuclear generation share
     for every available year.
  3. Merges in state_goals.json (hand-curated RPS/CES/100%-mandate status —
     edit that file directly if a state's law changes; this script does not
     rewrite it).
  4. Renders the self-contained artifact.html in this same folder.

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


def build_site_data(gen, goals):
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
            'cur': {'year': int(last_yr), 'f': cur['f'], 'r': cur['r'], 'n': cur['n'], 't': cur['t']},
            'd10': {'from_year': int(prior_yr), 'df': round(cur['f'] - prior['f'], 1), 'dr': round(cur['r'] - prior['r'], 1)}
        }

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

    last_data_year = int(years_all[-1])
    # EIA has historically released this workbook in September, covering the
    # prior calendar year, with the following year's edition around October.
    released = f"September {last_data_year + 1}"
    next_release = f"October {last_data_year + 2}"
    return {
        'meta': {'lastDataYear': last_data_year, 'released': released, 'nextRelease': next_release,
                 'source': 'EIA, Net Generation by State by Type of Producer by Energy Source (1990–present)'},
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
    site_data = build_site_data(gen, goals)
    out_path = render_html(site_data)
    print(f"Built {out_path} — data through {site_data['meta']['lastDataYear']}, {len(site_data['states'])} states.")


if __name__ == '__main__':
    main()
