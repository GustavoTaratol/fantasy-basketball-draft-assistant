from openpyxl import load_workbook
from pathlib import Path
from datetime import datetime
from io import StringIO
from urllib.request import Request, urlopen
import re
import json
import pandas as pd

SOURCE = Path('/Users/MSWTeam/Downloads/BBL_Draft_Assistant.xlsx')
OUTPUT = Path('/Users/MSWTeam/Documents/New project/Deliverables/BBL Draft Assistant/BBL_Draft_Assistant_Offline.html')

wbv = load_workbook(SOURCE, data_only=True, read_only=False)
db = wbv['PLAYER DATABASE']
workbook_players = {}
for r in range(2, db.max_row + 1):
    name = db.cell(r, 4).value
    if not name:
        continue
    record = {
        'rank': int(db.cell(r, 3).value or r - 1),
        'name': str(name),
        'pos': str(db.cell(r, 5).value or ''),
        'team': str(db.cell(r, 6).value or ''),
        'adp': float(db.cell(r, 7).value or 999),
        'value': float(db.cell(r, 8).value or 0),
        'playoff': float(db.cell(r, 9).value or 0),
    }
    workbook_players[re.sub(r'[^a-z0-9]', '', str(name).lower())] = record

cal = wbv['PLAYOFF CALENDAR']
dates = []
for c in range(2, 23):
    v = cal.cell(1, c).value
    if isinstance(v, datetime):
        dates.append(v.strftime('%Y-%m-%d'))
    else:
        dates.append(str(v))
schedules = {}
opponents = {}
for r in range(2, 32):
    team = cal.cell(r, 1).value
    if team:
        opponents[str(team)] = [str(cal.cell(r, c).value or '') for c in range(2, 23)]
        schedules[str(team)] = [bool(value) for value in opponents[str(team)]]
nba_games = [int(cal.cell(34, c).value or 0) for c in range(2, 23)]

# CBS Sports ADP is the authoritative player pool. The workbook supplies the
# custom BBL and playoff values for players it already contains.
cbs_url = 'https://www.cbssports.com/fantasy/basketball/draft/averages/'
snapshot = Path(__file__).with_name('cbs_adp_snapshot.csv')
if snapshot.exists() and len(pd.read_csv(snapshot)) >= 170:
    adp_table = pd.read_csv(snapshot)
else:
    req = Request(cbs_url, headers={'User-Agent': 'Mozilla/5.0'})
    html = urlopen(req, timeout=30).read().decode('utf-8', 'ignore')
    raw = pd.read_html(StringIO(html))[0]
    records = []
    for _, row in raw.iterrows():
        parts = [x.strip() for x in re.split(r'\s{2,}', str(row['Player']).strip())]
        hits = [i for i, value in enumerate(parts) if value in {'PG','SG','SF','PF','C'} and i > 0 and i + 1 < len(parts)]
        if hits:
            i = hits[-1]
            records.append({'rank': int(row['Rank']), 'name': parts[i-1], 'pos': parts[i], 'team': parts[i+1], 'adp': float(row['Avg Pos'])})
    adp_table = pd.DataFrame(records)
    adp_table.to_csv(snapshot, index=False)

team_alias = {'SA':'SAS','NY':'NYK','NO':'NOP','GS':'GSW','PHO':'PHX'}
players = []
for row in adp_table.to_dict('records'):
    key = re.sub(r'[^a-z0-9]', '', str(row['name']).lower())
    prior = workbook_players.get(key)
    team = team_alias.get(str(row['team']), str(row['team']))
    games = schedules.get(team, [False] * 21)
    low_volume_games = sum(1 for i, plays in enumerate(games) if plays and nba_games[i] <= 6)
    if prior:
        value, playoff = prior['value'], prior['playoff']
    else:
        value = round(max(40, 100 - (float(row['adp']) - 1) * .34), 1)
        playoff = round(min(100, 50 + sum(games) * 2 + low_volume_games * 1.5), 1)
    players.append({'rank': int(row['rank']), 'name': str(row['name']), 'pos': str(row['pos']), 'team': team, 'adp': float(row['adp']), 'value': value, 'playoff': playoff, 'estimated': False})

# CBS currently publishes only 177 players with a recorded ADP. Extend the
# draftable pool to 225 with players appearing in CBS's current Top 200 and
# roster-trends data. Their ordering is an estimate and is labeled as such.
deep_snapshot = Path(__file__).with_name('cbs_deep_pool_snapshot.csv')
if deep_snapshot.exists() and len(pd.read_csv(deep_snapshot)) >= 48:
    deep_table = pd.read_csv(deep_snapshot)
else:
    def fetch_table(url):
        req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        page = urlopen(req, timeout=30).read().decode('utf-8', 'ignore')
        return pd.read_html(StringIO(page))[0]
    rankings = fetch_table('https://www.cbssports.com/fantasy/basketball/rankings/roto/top200/')
    rank_order = {}
    for _, row in rankings.iterrows():
        label = re.sub(r'(PG|SG|SF|PF|C)$', '', str(row['Player']).strip())
        label = re.sub(r'\s+[A-Z]\.\s+', ' ', label)
        rank_order[re.sub(r'[^a-z0-9]', '', label.lower())] = int(row['RK'])
    trends = fetch_table('https://www.cbssports.com/fantasy/basketball/trends/added/all/')
    existing = {re.sub(r'[^a-z0-9]', '', p['name'].lower()) for p in players}
    candidates = []
    for _, row in trends.iterrows():
        parts = [x.strip() for x in re.split(r'\s{2,}', str(row['Player']).strip())]
        hits = [i for i, value in enumerate(parts) if value in {'PG','SG','SF','PF','C'} and i > 0 and i + 1 < len(parts)]
        if not hits:
            continue
        i = hits[-1]
        name, pos, team = parts[i-1], parts[i], parts[i+1]
        key = re.sub(r'[^a-z0-9]', '', name.lower())
        if key in existing:
            continue
        candidates.append({'name': name, 'pos': pos, 'team': team, 'rostered': float(row['Current Week']), 'cbsRank': rank_order.get(key, 999)})
    candidates.sort(key=lambda x: (x['cbsRank'] >= 999, x['cbsRank'], -x['rostered'], x['name']))
    deep_table = pd.DataFrame(candidates[:48])
    deep_table.to_csv(deep_snapshot, index=False)

for offset, row in enumerate(deep_table.to_dict('records'), start=178):
    team = team_alias.get(str(row['team']), str(row['team']))
    games = schedules.get(team, [False] * 21)
    low_volume_games = sum(1 for i, plays in enumerate(games) if plays and nba_games[i] <= 6)
    adp = float(offset)
    players.append({'rank': offset, 'name': str(row['name']), 'pos': str(row['pos']), 'team': team, 'adp': adp, 'value': round(max(30, 100 - (adp - 1) * .34), 1), 'playoff': round(min(100, 50 + sum(games) * 2 + low_volume_games * 1.5), 1), 'estimated': True})

# Rebuild Player Value from CBS season projections and the BBL scoring system.
# CBS projections omit personal fouls, so PF is conservatively estimated by
# listed position and disclosed in the app.
projection_snapshot = Path(__file__).with_name('cbs_projection_snapshot.csv')
required_projection_cols = {'name','gp','fgm','fga','ftm','fta','threepm','pts','reb','ast','stl','turnovers','blk'}
if projection_snapshot.exists() and required_projection_cols.issubset(pd.read_csv(projection_snapshot, nrows=1).columns):
    projection_table = pd.read_csv(projection_snapshot)
else:
    projection_rows = {}
    for position in ['PG','SG','SF','PF','C']:
        url = f'https://www.cbssports.com/fantasy/basketball/stats/{position}/2026/season/projections/'
        req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        page = urlopen(req, timeout=30).read().decode('utf-8', 'ignore')
        table = pd.read_html(StringIO(page))[0]
        for _, row in table.iterrows():
            parts = [x.strip() for x in re.split(r'\s{2,}', str(row['Player']).strip())]
            hits = [i for i, value in enumerate(parts) if value in {'PG','SG','SF','PF','C'} and i > 0]
            if not hits:
                continue
            i = hits[-1]
            name = parts[i-1]
            key = re.sub(r'[^a-z0-9]', '', name.lower())
            projection_rows[key] = {
                'name': name,
                'gp': float(row['gp  Games Played']),
                'fgm': float(row['fgm  Field Goals Made']),
                'fga': float(row['fga  Field Goal Attempts']),
                'ftm': float(row['ftm  Free Throws Made']),
                'fta': float(row['fta  Free Throw Attempts']),
                'threepm': float(row['3pm  Three-point Field Goals Made']),
                'pts': float(row['pts  Points']),
                'reb': float(row['reb  Total Rebounds']),
                'ast': float(row['ast  Assists']),
                'stl': float(row['stl  Steals']),
                'turnovers': float(row['to  Turnovers']),
                'blk': float(row['blk  Blocks']),
            }
    projection_table = pd.DataFrame(projection_rows.values())
    projection_table.to_csv(projection_snapshot, index=False)

projection_by_name = {re.sub(r'[^a-z0-9]', '', str(row['name']).lower()): row for row in projection_table.to_dict('records')}
pf_per_game = {'PG': 2.0, 'SG': 2.1, 'SF': 2.2, 'PF': 2.4, 'C': 2.6}
matched_adp, matched_ppg = [], []
for player in players:
    row = projection_by_name.get(re.sub(r'[^a-z0-9]', '', player['name'].lower()))
    if not row or not row['gp']:
        continue
    special = row['reb'] if player['pos'] in {'PG','SG'} else row['ast'] if player['pos'] in {'SF','PF'} else row['threepm'] if player['pos'] == 'C' else 0
    total = (row['ast']*2 + row['blk']*4 + row['fgm']*2 - row['fga'] + row['ftm'] - row['fta'] + row['pts'] + row['stl']*4 - row['turnovers']*2 + row['reb'] + special - pf_per_game.get(player['pos'],2.2)*row['gp'])
    player['bblPpg'] = round(total / row['gp'], 2)
    player['projectionEstimated'] = False
    matched_adp.append(player['adp'])
    matched_ppg.append(player['bblPpg'])

paired = sorted(zip(matched_adp, matched_ppg))
for player in players:
    if 'bblPpg' in player:
        continue
    lower = max((x for x in paired if x[0] <= player['adp']), default=paired[0])
    upper = min((x for x in paired if x[0] >= player['adp']), default=paired[-1])
    if upper[0] == lower[0]:
        estimate = lower[1]
    else:
        estimate = lower[1] + (upper[1]-lower[1]) * (player['adp']-lower[0]) / (upper[0]-lower[0])
    player['bblPpg'] = round(estimate, 2)
    player['projectionEstimated'] = True

ppg_values = sorted(p['bblPpg'] for p in players)
high = max(ppg_values)
for player in players:
    player['value'] = round(max(0, min(100, player['bblPpg']/high*100)), 1)

# Schedule score rewards total playoff games and low-volume dates, using the
# supplied actual-date calendar. It remains secondary to projected BBL value.
schedule_raw = {}
for team, games in schedules.items():
    schedule_raw[team] = sum((1.25 if nba_games[i] <= 5 else 1.12 if nba_games[i] <= 7 else 1.0 if nba_games[i] <= 9 else .88) for i, plays in enumerate(games) if plays)
schedule_min, schedule_max = min(schedule_raw.values()), max(schedule_raw.values())
for player in players:
    raw = schedule_raw.get(player['team'], schedule_min)
    player['playoff'] = round(45 + (raw-schedule_min)/(schedule_max-schedule_min)*55, 1)

settings = wbv['SETTINGS']
seed = {
    'players': players,
    'dates': dates,
    'schedules': schedules,
    'opponents': opponents,
    'nbaGames': nba_games,
    'settings': {
        'teams': int(settings['B2'].value or 10),
        'format': str(settings['B3'].value or 'Third Round Reversal (3RR)'),
        'slot': int(settings['B4'].value or 8),
        'rounds': 13,
        'maxActive': int(settings['B6'].value or 5),
        'weights': [float(settings['B7'].value or .7), float(settings['B8'].value or .2), float(settings['B9'].value or .1)],
        'picks': [8,13,23,38,43,58,63,78,83,98,103,118,123],
        'source': 'CBS Sports NBA ADP (1–177) with CBS-ranked deep pool estimates (178–225)',
        'sourceUrl': cbs_url,
        'sourceDate': '2026-10-07',
    }
}

template = r'''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>BBL Draft Assistant</title>
<style>
:root{--bg:#090c10;--panel:#11161c;--panel2:#171d24;--line:#29313b;--text:#f5f7f9;--muted:#919ca8;--orange:#ff6b2c;--orange2:#ff9b66;--green:#40d69b;--red:#ff626e;--amber:#ffc857;--blue:#69a9ff;--radius:14px}*{box-sizing:border-box}html{background:var(--bg)}body{margin:0;color:var(--text);background:radial-gradient(circle at 80% -10%,#352011 0,transparent 30%),var(--bg);font-family:Inter,ui-sans-serif,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;font-size:14px}button,input,select{font:inherit}.app{min-height:100vh}.topbar{height:70px;border-bottom:1px solid var(--line);display:flex;align-items:center;gap:28px;padding:0 26px;background:rgba(9,12,16,.94);position:sticky;top:0;z-index:20;backdrop-filter:blur(12px)}.brand{font-size:18px;font-weight:900;letter-spacing:-.03em}.brand b{color:var(--orange)}.live{color:var(--orange);font-size:11px;font-weight:900;letter-spacing:.12em}.nav{display:flex;gap:6px;margin-left:auto}.nav button{background:transparent;color:var(--muted);border:0;padding:9px 12px;border-radius:9px;cursor:pointer;font-weight:700}.nav button.active,.nav button:hover{color:white;background:var(--panel2)}.shell{max-width:1500px;margin:auto;padding:24px}.hero{display:grid;grid-template-columns:minmax(0,1fr) 210px 210px;gap:12px;margin-bottom:18px}.hero-main,.pick-card{background:linear-gradient(145deg,var(--panel2),var(--panel));border:1px solid var(--line);border-radius:var(--radius);padding:20px}.eyebrow{font-size:11px;color:var(--orange2);font-weight:900;letter-spacing:.13em;text-transform:uppercase}.hero h1{font-size:27px;margin:7px 0 6px;letter-spacing:-.04em}.hero p{margin:0;color:var(--muted)}.pick-card span{display:block;color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.1em}.pick-card strong{display:block;font-size:31px;margin-top:5px}.layout{display:grid;grid-template-columns:minmax(0,1fr) 330px;gap:18px}.panel{background:var(--panel);border:1px solid var(--line);border-radius:var(--radius);overflow:hidden}.panel-head{padding:16px 18px;border-bottom:1px solid var(--line);display:flex;align-items:center;gap:12px}.panel-head h2{font-size:15px;margin:0}.count{color:var(--muted);font-size:12px}.tools{display:flex;gap:8px;margin-left:auto}.search,.select,.number{background:#0c1015;border:1px solid var(--line);color:white;border-radius:9px;padding:9px 11px;outline:none}.search{width:220px}.search:focus,.select:focus,.number:focus{border-color:var(--orange)}.select{max-width:110px}.table-wrap{overflow:auto;max-height:680px}table{width:100%;border-collapse:collapse;white-space:nowrap}th{position:sticky;top:0;background:#131920;color:var(--muted);text-align:left;font-size:10px;letter-spacing:.08em;text-transform:uppercase;padding:11px 10px;border-bottom:1px solid var(--line);z-index:2}td{padding:10px;border-bottom:1px solid #202730}tbody tr:hover{background:#171e26}.rank{color:var(--muted);width:42px}.player strong{display:block}.player small{color:var(--muted)}.score{font-weight:900;color:white}.score.high{color:var(--green)}.pill{display:inline-flex;padding:4px 7px;border-radius:999px;font-size:10px;font-weight:900;letter-spacing:.04em}.take{background:#153f32;color:#65e4b4}.good{background:#253653;color:#9fc5ff}.wait{background:#3b3020;color:#ffd686}.actions{display:flex;gap:5px}.action{border:1px solid var(--line);background:#1a2129;color:white;padding:6px 9px;border-radius:7px;font-weight:800;font-size:11px;cursor:pointer}.action:hover{border-color:var(--orange)}.action.mine{background:var(--orange);border-color:var(--orange);color:#160903}.action.taken{color:var(--muted)}.side{display:grid;gap:18px;align-content:start}.recommend{padding:18px;background:linear-gradient(145deg,#25170f,#15181c);border-color:#6c371e}.recommend .name{font-size:23px;font-weight:900;letter-spacing:-.04em;margin:8px 0 3px}.recommend .meta{color:var(--muted)}.recommend .why{margin-top:14px;color:#d8dee4;line-height:1.5}.side-list{padding:8px 16px 14px}.side-row{display:flex;justify-content:space-between;gap:10px;padding:10px 0;border-bottom:1px solid #222a33}.side-row:last-child{border:0}.side-row small{color:var(--muted)}.empty{padding:28px 18px;text-align:center;color:var(--muted)}.btn-row{display:flex;gap:8px;padding:14px 16px;border-top:1px solid var(--line)}.btn{flex:1;padding:9px;border-radius:8px;border:1px solid var(--line);background:#161d24;color:white;cursor:pointer;font-weight:800}.btn:hover{border-color:var(--orange)}.alert{display:grid;grid-template-columns:68px 1fr auto;gap:10px;align-items:center;padding:11px 0;border-bottom:1px solid #222a33}.alert:last-child{border:0}.critical{color:var(--red);font-size:10px;font-weight:900}.open{color:var(--amber);font-size:10px;font-weight:900}.date{font-weight:800}.slots{color:var(--muted);font-size:12px}.view{display:none}.view.active{display:block}.full-grid{display:grid;grid-template-columns:1fr 1fr;gap:18px}.roster-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:10px;padding:16px}.roster-card{background:var(--panel2);border:1px solid var(--line);padding:14px;border-radius:10px;display:flex;justify-content:space-between;align-items:center}.roster-card small{color:var(--muted)}.coverage{padding:18px;overflow:auto}.heat{min-width:920px;display:grid;grid-template-columns:130px repeat(21,1fr);gap:4px}.heat>div{min-height:34px;display:flex;align-items:center;justify-content:center;border-radius:5px;background:#151b22;font-size:10px}.heat .label{justify-content:flex-start;padding-left:8px;font-weight:800}.heat .head{color:var(--muted);writing-mode:vertical-rl;height:78px;justify-content:end;padding-bottom:6px}.heat .full{background:#154333;color:#75e5b9}.heat .mid{background:#4b3618;color:#ffd27b}.heat .low{background:#4c2028;color:#ff9da5}.settings{padding:20px;display:grid;grid-template-columns:repeat(2,minmax(240px,1fr));gap:18px}.field{display:grid;gap:7px}.field label{color:var(--muted);font-weight:700}.field input{width:100%}.range-line{display:flex;gap:12px;align-items:center}.range-line input{flex:1}.range-line output{width:42px;text-align:right;font-weight:900}.settings-note{grid-column:1/-1;padding:14px;border-radius:10px;background:#141a21;color:var(--muted);line-height:1.5}.footer{color:#5e6873;text-align:center;padding:25px}.toast{position:fixed;right:20px;bottom:20px;background:white;color:#111;padding:12px 16px;border-radius:9px;font-weight:800;transform:translateY(90px);opacity:0;transition:.2s;z-index:50}.toast.show{transform:none;opacity:1}@media(max-width:1000px){.layout{grid-template-columns:1fr}.side{grid-template-columns:1fr 1fr}.hero{grid-template-columns:1fr 150px 150px}}@media(max-width:720px){.topbar{height:auto;min-height:62px;padding:12px 14px;flex-wrap:wrap;gap:10px}.nav{order:3;width:100%;overflow:auto}.shell{padding:14px}.hero{grid-template-columns:1fr 1fr}.hero-main{grid-column:1/-1}.layout,.full-grid,.side,.settings{grid-template-columns:1fr}.tools{width:100%;margin:0;flex-wrap:wrap}.panel-head{flex-wrap:wrap}.search{width:100%}.hide-mobile{display:none}.table-wrap{max-height:none}.settings-note{grid-column:auto}.pick-card{padding:14px}.pick-card strong{font-size:25px}}
</style>
</head>
<body>
<div class="app">
  <header class="topbar"><div class="brand"><b>BBL</b> Draft Assistant</div><div class="live">● OFFLINE MODE</div><nav class="nav"><button class="active" data-view="draft">Draft Room</button><button data-view="roster">My Roster</button><button data-view="coverage">Playoff Coverage</button><button data-view="settings">Settings</button></nav></header>
  <main class="shell">
    <section id="draft" class="view active">
      <div class="hero"><div class="hero-main"><div class="eyebrow">Live recommendation engine</div><h1>Build value now. Win the playoff schedule later.</h1><p id="heroCopy">Rankings recalculate after every manual pick.</p></div><div class="pick-card"><span>Your current pick</span><strong id="currentPick">8</strong></div><div class="pick-card"><span>Your next pick</span><strong id="nextPick">13</strong></div></div>
      <div class="layout"><section class="panel"><div class="panel-head"><h2>Available players</h2><span class="count" id="availableCount"></span><div class="tools"><input id="search" class="search" placeholder="Search player or team"><select id="position" class="select"><option value="">All positions</option><option>PG</option><option>SG</option><option>SF</option><option>PF</option><option>C</option></select></div></div><div class="table-wrap"><table><thead><tr><th>#</th><th>Player</th><th>ADP</th><th>BBL</th><th>Playoffs</th><th>Fit</th><th>Overall</th><th>Call</th><th>Mark pick</th></tr></thead><tbody id="playerRows"></tbody></table></div></section>
      <aside class="side"><section class="panel recommend"><div class="eyebrow">Best available</div><div class="name" id="bestName">—</div><div class="meta" id="bestMeta"></div><div class="why" id="bestWhy"></div></section><section class="panel"><div class="panel-head"><h2>My roster</h2><span class="count" id="rosterCount"></span></div><div id="miniRoster" class="side-list"></div><div class="btn-row"><button id="undo" class="btn">Undo</button><button id="reset" class="btn">Reset draft</button></div></section><section class="panel"><div class="panel-head"><h2>Coverage alerts</h2></div><div id="miniAlerts" class="side-list"></div></section></aside></div>
    </section>
    <section id="roster" class="view"><div class="hero"><div class="hero-main"><div class="eyebrow">Roster construction</div><h1>My roster</h1><p>Every player marked Mine appears here and contributes to schedule coverage.</p></div><div class="pick-card"><span>Players drafted</span><strong id="rosterTotal">0</strong></div><div class="pick-card"><span>Rounds remaining</span><strong id="roundsRemaining">13</strong></div></div><section class="panel"><div class="panel-head"><h2>Drafted players</h2></div><div id="rosterGrid" class="roster-grid"></div></section></section>
    <section id="coverage" class="view"><div class="hero"><div class="hero-main"><div class="eyebrow">Three playoff periods</div><h1>Daily lineup coverage</h1><p>Green dates are full. Red dates have three or more unused active slots.</p></div><div class="pick-card"><span>Under-covered dates</span><strong id="gapCount">21</strong></div><div class="pick-card"><span>Unused active slots</span><strong id="unusedSlots">105</strong></div></div><section class="panel"><div class="panel-head"><h2>Playoff schedule heatmap</h2></div><div class="coverage"><div id="heatmap" class="heat"></div></div></section></section>
    <section id="settings" class="view"><div class="hero"><div class="hero-main"><div class="eyebrow">Recommendation controls</div><h1>Draft settings</h1><p>Adjust how the assistant balances rankings, playoff strength, and roster fit.</p></div></div><section class="panel"><div class="settings"><div class="field"><label>Maximum active players per day</label><input id="maxActive" class="number" type="number" min="1" max="15"></div><div class="field"><label>Draft label</label><input class="number" value="10-team • 13 rounds • 3RR" disabled></div><div class="field"><label>Player value weight</label><div class="range-line"><input id="weightValue" type="range" min="0" max="100"><output id="outValue"></output></div></div><div class="field"><label>Playoff value weight</label><div class="range-line"><input id="weightPlayoff" type="range" min="0" max="100"><output id="outPlayoff"></output></div></div><div class="field"><label>Roster fit weight</label><div class="range-line"><input id="weightFit" type="range" min="0" max="100"><output id="outFit"></output></div></div><div class="settings-note">Weights are normalized automatically, so they do not need to add to 100. Draft progress and settings are saved only in this browser on this device.</div></div></section></section>
  </main><div class="footer">BBL Draft Assistant · Rankings and schedule embedded from the supplied workbook</div>
</div><div id="toast" class="toast"></div>
<script>
const SEED=__SEED__;
const STORE='bbl-draft-assistant-v1';
let state={mine:[],taken:[],history:[],maxActive:SEED.settings.maxActive,weights:SEED.settings.weights};
try{const saved=JSON.parse(localStorage.getItem(STORE));if(saved)state={...state,...saved}}catch(e){}
const $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
const save=()=>localStorage.setItem(STORE,JSON.stringify(state));
const esc=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const drafted=()=>SEED.players.filter(p=>state.mine.includes(p.name));
const scheduleLoads=()=>SEED.dates.map((_,i)=>drafted().reduce((n,p)=>n+(SEED.schedules[p.team]?.[i]?1:0),0));
function fitScore(p){if(!state.mine.length)return 70;const loads=scheduleLoads(),games=SEED.schedules[p.team]||[];let useful=0,total=0;games.forEach((g,i)=>{if(g){total++;const room=Math.max(0,state.maxActive-loads[i]);if(room)useful+=1+(SEED.nbaGames[i]<=5?.35:0)+(room>=3?.2:0)}});const schedule=total?Math.min(100,useful/Math.max(1,total)*78):45;const posCount=drafted().filter(x=>x.pos===p.pos).length;const need=Math.max(35,100-posCount*22);return Math.round(schedule*.72+need*.28)}
function scoredPlayers(){const sum=state.weights.reduce((a,b)=>a+b,0)||1;return SEED.players.map(p=>{const fit=fitScore(p);const overall=(p.value*state.weights[0]+p.playoff*state.weights[1]+fit*state.weights[2])/sum;const target=SEED.settings.picks[state.mine.length]||999;const call=p.adp<=target+3?'TAKE NOW':p.adp>=((SEED.settings.picks[state.mine.length+1]||target+15)+3)?'WAIT':'GOOD VALUE';return {...p,fit,overall,call}}).filter(p=>!state.mine.includes(p.name)&&!state.taken.includes(p.name)).sort((a,b)=>b.overall-a.overall||a.rank-b.rank)}
function mark(name,type){if(state.mine.includes(name)||state.taken.includes(name))return;state[type].push(name);state.history.push({name,type});save();render();toast(type==='mine'?'Added to your roster':'Marked as taken')}
function undo(){const last=state.history.pop();if(!last)return toast('Nothing to undo');state[last.type]=state[last.type].filter(n=>n!==last.name);save();render();toast('Last pick undone')}
function resetDraft(){if(!confirm('Reset every Mine and Taken selection?'))return;state.mine=[];state.taken=[];state.history=[];save();render();toast('Draft reset')}
function toast(msg){const t=$('#toast');t.textContent=msg;t.classList.add('show');setTimeout(()=>t.classList.remove('show'),1500)}
function renderDraft(){let list=scoredPlayers();const q=$('#search').value.trim().toLowerCase(),pos=$('#position').value;if(q)list=list.filter(p=>(p.name+' '+p.team).toLowerCase().includes(q));if(pos)list=list.filter(p=>p.pos===pos);$('#availableCount').textContent=`${list.length} available`;$('#playerRows').innerHTML=list.map((p,i)=>`<tr><td class="rank">${i+1}</td><td class="player"><strong>${esc(p.name)}</strong><small>${esc(p.team)} · ${esc(p.pos)} · CBS #${p.rank}</small></td><td>${p.adp.toFixed(1)}</td><td>${p.value.toFixed(1)}</td><td>${p.playoff.toFixed(1)}</td><td>${p.fit}</td><td class="score ${p.overall>=90?'high':''}">${p.overall.toFixed(1)}</td><td><span class="pill ${p.call==='TAKE NOW'?'take':p.call==='WAIT'?'wait':'good'}">${p.call}</span></td><td><div class="actions"><button class="action mine" data-mine="${esc(p.name)}">Mine</button><button class="action taken" data-taken="${esc(p.name)}">Taken</button></div></td></tr>`).join('');const all=scoredPlayers(),best=all[0];if(best){$('#bestName').textContent=best.name;$('#bestMeta').textContent=`${best.team} · ${best.pos} · ${best.overall.toFixed(1)} overall`;$('#bestWhy').textContent=`BBL value ${best.value.toFixed(1)}, playoff value ${best.playoff.toFixed(1)}, and ${best.fit}% roster fit make this the strongest available combination.`}const current=SEED.settings.picks[state.mine.length]||'—',next=SEED.settings.picks[state.mine.length+1]||'—';$('#currentPick').textContent=current;$('#nextPick').textContent=next;$('#heroCopy').textContent=`${state.mine.length} of ${SEED.settings.rounds} roster spots filled · ${state.taken.length} opponents' picks recorded`;const mine=drafted();$('#rosterCount').textContent=`${mine.length}/${SEED.settings.rounds}`;$('#miniRoster').innerHTML=mine.length?mine.slice(-6).reverse().map(p=>`<div class="side-row"><div><b>${esc(p.name)}</b><br><small>${p.team} · ${p.pos}</small></div><b>${p.value.toFixed(1)}</b></div>`).join(''):'<div class="empty">Use Mine beside a player to start your roster.</div>';renderAlerts()}
function renderAlerts(){const loads=scheduleLoads();const gaps=SEED.dates.map((date,i)=>({date,i,load:loads[i],open:Math.max(0,state.maxActive-loads[i]),games:SEED.nbaGames[i]})).filter(x=>x.open>0).sort((a,b)=>b.open-a.open||a.games-b.games);$('#miniAlerts').innerHTML=gaps.slice(0,5).map(g=>`<div class="alert"><div class="${g.open>=3?'critical':'open'}">${g.open>=3?'CRITICAL':'OPEN'}</div><div><div class="date">${new Date(g.date+'T12:00:00').toLocaleDateString(undefined,{month:'short',day:'numeric'})}</div><div class="slots">${g.games} NBA games</div></div><b>${g.open} slots</b></div>`).join('')||'<div class="empty">Every playoff date is fully covered.</div>';$('#gapCount').textContent=gaps.length;$('#unusedSlots').textContent=gaps.reduce((s,g)=>s+g.open,0)}
function renderRoster(){const mine=drafted();$('#rosterTotal').textContent=mine.length;$('#roundsRemaining').textContent=Math.max(0,SEED.settings.rounds-mine.length);$('#rosterGrid').innerHTML=mine.length?mine.map(p=>`<div class="roster-card"><div><b>${esc(p.name)}</b><br><small>${p.team} · ${p.pos} · ADP ${p.adp.toFixed(1)}</small></div><button class="action" data-remove="${esc(p.name)}">Remove</button></div>`).join(''):'<div class="empty">No players drafted yet.</div>'}
function renderCoverage(){const loads=scheduleLoads();let html='<div class="label">Period</div>'+SEED.dates.map((d,i)=>`<div class="head">P${i<7?22:i<14?23:24} · ${new Date(d+'T12:00:00').toLocaleDateString(undefined,{month:'short',day:'numeric'})}</div>`).join('');html+='<div class="label">My players</div>'+loads.map(n=>`<div class="${n>=state.maxActive?'full':n>=state.maxActive-2?'mid':'low'}">${n}/${state.maxActive}</div>`).join('');html+='<div class="label">NBA games</div>'+SEED.nbaGames.map(n=>`<div>${n}</div>`).join('');drafted().forEach(p=>{html+=`<div class="label">${esc(p.name.split(' ').slice(-1)[0])} · ${p.team}</div>`+(SEED.schedules[p.team]||[]).map(g=>`<div class="${g?'full':''}">${g?'●':''}</div>`).join('')});$('#heatmap').innerHTML=html}
function renderSettings(){$('#maxActive').value=state.maxActive;['Value','Playoff','Fit'].forEach((n,i)=>{$('#weight'+n).value=Math.round(state.weights[i]*100);$('#out'+n).textContent=Math.round(state.weights[i]*100)+'%'})}
function render(){renderDraft();renderRoster();renderCoverage();renderSettings()}
document.addEventListener('click',e=>{const b=e.target.closest('button');if(!b)return;if(b.dataset.view){$$('.nav button').forEach(x=>x.classList.toggle('active',x===b));$$('.view').forEach(x=>x.classList.toggle('active',x.id===b.dataset.view));return}if(b.dataset.mine)mark(b.dataset.mine,'mine');if(b.dataset.taken)mark(b.dataset.taken,'taken');if(b.dataset.remove){state.mine=state.mine.filter(n=>n!==b.dataset.remove);state.history=state.history.filter(x=>x.name!==b.dataset.remove);save();render()} });
$('#search').addEventListener('input',renderDraft);$('#position').addEventListener('change',renderDraft);$('#undo').addEventListener('click',undo);$('#reset').addEventListener('click',resetDraft);$('#maxActive').addEventListener('change',e=>{state.maxActive=Math.max(1,Number(e.target.value)||5);save();render()});['Value','Playoff','Fit'].forEach((n,i)=>$('#weight'+n).addEventListener('input',e=>{state.weights[i]=Number(e.target.value)/100;save();render()}));render();
</script>
</body></html>'''

project_dir = Path(__file__).parent
extra_css = (project_dir / 'v2_extra.css').read_text(encoding='utf-8')
additions = (project_dir / 'v2_additions.html').read_text(encoding='utf-8')
app_js = (project_dir / 'v2_app.js').read_text(encoding='utf-8').replace('__SEED__', json.dumps(seed, separators=(',', ':')))
template = template.replace('</style>', extra_css + '\n</style>')
template = template.replace('</main>', additions + '\n</main>')
template = re.sub(r'<script>.*?</script>', lambda _: '<script>\n' + app_js + '\n</script>', template, flags=re.S)
OUTPUT.parent.mkdir(parents=True, exist_ok=True)
OUTPUT.write_text(template, encoding='utf-8')
print(OUTPUT)
