# State Power Transition Tracker

Cowork artifact id: `state-power-transition-tracker`

Tracks, for all 50 states + DC: the fossil-vs-renewable-vs-nuclear share of
electricity generation over time (source: EIA), paired with each state's
carbon/clean-energy goal status (RPS, CES, or a binding 100% mandate).

## Files
- `build_tracker.py` — full pipeline: downloads the EIA workbook, recomputes
  shares, merges `state_goals.json`, renders `artifact.html`.
- `template.html` — the artifact's HTML/CSS/JS shell, with a `__DATA_JSON__`
  placeholder `build_tracker.py` fills in.
- `state_goals.json` — hand-curated table of each state's RPS/CES/100%-mandate
  status. Not auto-derived — edit this file directly when a state's law
  changes, then rerun the pipeline.
- `gen_state.xlsx` — the most recently downloaded EIA workbook (regenerated
  each run, safe to delete).

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
- EIA's state generation workbook updates roughly annually — historically
  released in September, covering the prior year, with the following year's
  release around October.
- Fossil = coal, natural gas, petroleum, other gases. Renewable = conventional
  hydroelectric, wind, solar thermal/PV, geothermal, wood and other biomass.
  Nuclear is tracked separately as carbon-free but not renewable.
- Goal status was compiled July 2026 from NCSL's state RPS/CES brief, cross-
  checked against DSIRE and Environment America's 100%-clean-electricity
  tracking.
