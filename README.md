# BBL Draft Assistant

Offline NBA fantasy draft assistant built from the supplied BBL workbook and
CBS player data snapshots.

## Finished application

The distributable application is a single standalone HTML file:

`Deliverables/BBL Draft Assistant/BBL_Draft_Assistant_Offline.html`

It runs locally in a modern browser and does not require an internet
connection.

The same generated application is written to `docs/index.html` for GitHub
Pages. Configure Pages to deploy from the `main` branch and `/docs` folder.

## League configuration

- 10 teams
- 13 rounds
- Third Round Reversal (3RR)
- Daily starters: 1 G, 1 F, 1 C, and 2 G/F/C flex
- 8 bench spots
- Draft position can be changed from Settings

## Offline application source

- `build_offline.py` — reads the workbook and CBS snapshots and builds the
  standalone HTML application
- `v2_app.js` — application state, draft projections, scoring, lineup coverage,
  settings, and interactions
- `v2_extra.css` — application styling and responsive layout
- `v2_additions.html` — additional application panels and controls
- `cbs_adp_snapshot.csv` — CBS ADP snapshot
- `cbs_deep_pool_snapshot.csv` — extended CBS player pool
- `cbs_projection_snapshot.csv` — CBS projection snapshot

The workbook path and output path are configured near the top of
`build_offline.py`.

## Rebuilding

Python dependencies used by the generator:

- `openpyxl`
- `pandas`

Run:

```bash
python3 build_offline.py
```

The checked-in CBS snapshots allow the embedded player pool to be rebuilt
without downloading fresh CBS data.

## Data and saves

The generated HTML embeds the player data, projections, schedules, CSS, and
JavaScript. Draft progress is saved in the browser's local storage. Users can
also download and restore backup files from inside the application.
