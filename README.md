# State Power Transition Tracker

Cowork artifact id: `state-power-transition-tracker`

Tracks, for all 50 states + DC: the fossil-vs-renewable-vs-nuclear share of
electricity generation over time (source: EIA), paired with each state's
carbon/clean-energy goal status (RPS, CES, or a binding 100% mandate).

## Files
- `build_tracker.py` — full pipeline: downloads the EIA annual and monthly
  workbooks, recomputes shares, forecasts 2026-2032, merges `state_goals.json`,
  renders `artifact.html`.
- `template.html` — the artifact's HTML/CSS/JS shell, with a `__DATA_JSON__`
  placeholder `build_tracker.py` fills in.
- `state_goals.json` — hand-curated table of each state's RPS/CES/100%-mandate
  status. Not auto-derived — edit this file directly when a state's law
  changes, then rerun the pipeline. Used only for the on-card label; see
  "Forecast methodology" below for why it never touches the forecast itself.
- `gen_state.xlsx` — the most recently downloaded EIA annual workbook
  (regenerated each run, safe to delete).
- `gen_state_monthly.xlsx` — the most recently downloaded EIA monthly
  workbook, used to build a real actual for any year past the annual
  release (regenerated each run, safe to delete).

## Refresh procedure
1. `cd` into this folder and run `python3 build_tracker.py`.
2. Read the resulting `artifact.html` back to confirm it built cleanly.
3. Copy it to the scratch outputs directory (or reference it directly if the
   tool allows), then call `mcp__cowork__update_artifact` with
   `id: "state-power-transition-tracker"` and the new `html_path`.
4. Before rebuilding, do a quick web search for recent state RPS/CES/clean-
   energy legislation (e.g. "state renewable portfolio standard clean energy
   law 2026 change"). If a state's status changed, edit its entry in
   `state_goals.json` first.

## Data notes
- EIA's annual "Final" workbook updates roughly annually — historically
  released in September, covering the prior year, with the following year's
  release around October. Any full calendar year after that gets rolled up
  from EIA's monthly workbook instead (see `process_monthly_rollup()`) and
  is a real actual, just flagged `prelim` until EIA's own annual release
  catches up to it.
- Fossil = coal, natural gas, petroleum, other gases. Renewable = conventional
  hydroelectric, wind, solar thermal/PV, geothermal, wood and other biomass.
  Nuclear is tracked separately as carbon-free but not renewable.
- Goal status was compiled July 2026 from NCSL's state RPS/CES brief, cross-
  checked against DSIRE and Environment America's 100%-clean-electricity
  tracking.

## Forecast methodology (2026-2032)

Each state's renewable-share forecast is a **logistic (S-curve) fit to that
state's own actual history** — `renewable_share(t) = L / (1 + e^(-k(t-t0)))`,
fit with `scipy.optimize.curve_fit` against every actual year available
(2001-present). `L` (the ceiling), `k` (steepness), and `t0` (inflection
year) all come out of the state's own trajectory.

**`state_goals.json` never touches the fit.** No forecast function in
`build_tracker.py` accepts `goals` as an argument — structurally, a state's
legal mandate cannot bend the curve. `verify_goals_isolation()` proves this
on every build: it reruns the entire forecast with every state's goal
swapped for a fabricated 100%-by-2026 mandate and asserts the output is
byte-identical. If a state's real trajectory implies it won't hit its own
mandate — Oregon is the clearest example, its renewable share has been flat
to declining for a decade against a 100%-by-2040 mandate — the forecast says
so. That gap between the mandate and the modeled trajectory is the point of
this tracker, not a defect to paper over.

**Fit-quality gate.** A logistic fit is only used if it converges with
R² ≥ 0.75 (`MIN_LOGISTIC_R2`). States with a flat, noisy, or non-monotonic
history — a hydro-heavy state's year-to-year swings, or a state that just
hasn't started ramping yet — fail that gate and fall back to a bounded
linear projection anchored at the last actual value, flagged
`lowConfidence: true`. As of this writing, 19 of 51 states get a converged
S-curve; 32 fall back to the conservative linear projection.

**Ceiling-identifiability edge case.** A state still deep in its rise, with
no plateau yet visible in the actual data, is mathematically
under-determined — a logistic curve can't distinguish "will level off at
55%" from "will level off at 100%" without having seen the bend yet, so the
optimizer drifts to the edge of its search bound (100%) even though that
number isn't really pinned down by the data. `build_state_forecast()` flags
this as `ceilingUnresolved: true` when the fitted ceiling lands at ≥99.9%,
and the summary text says so explicitly rather than presenting an
unsupported round number with false confidence. The near-term forecast
values (through 2032) are still the same R²-validated extrapolation of the
recent trend either way — only the long-run ceiling claim is unreliable in
this case.

**Nuclear and fossil are dependents, not independently fit.** Nuclear isn't
undergoing a diffusion transition in most states, so an S-curve is the wrong
shape for it — it gets a damped (30%-strength) continuation of its own
10-year trend instead. Fossil is always the remainder: `100 - renewable -
nuclear`, clipped at 0, with renewable yielding to nuclear if the two would
ever sum past 100 (this has not actually triggered for any state to date).

**Total generation** is projected per state from its own 10-year CAGR,
capped at ±5%/year so a low-base state (a single plant coming online or
retiring can look like a huge percentage swing) can't compound into an
implausible 2032 total. The national chart is a pure weighted rollup of the
51 state forecasts — there is no separate national-level fit — and
`verify_national_reconciliation()` proves that structurally on every build
by recomputing the national line independently and asserting it matches.

**Every build runs four guards before rendering:** `verify_goals_isolation`,
`verify_forecast_bounds` (every projected state-year's three shares sum to
100, each within [0, 100]), `verify_total_generation` (every projected total
stays positive), and `verify_national_reconciliation`. Any of the four
raising an exception fails the build rather than shipping bad numbers.
