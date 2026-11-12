#!/usr/bin/env python3
"""Rebuild data/schedule.json from the 'Seat Layout - Latest' sheet of the RTO workbook.

Usage:  python tools/import_seat_layout.py "Seat-FY27_Jan-Apr 1.xlsx" [--sheet "Seat Layout - Latest"] [--dry-run]

Sheet format: each seat block has a header row of seat IDs (e.g. 21.66E) followed by
five rows MON..FRI; each cell is "Team\nNickname", FREE, or RESERVED: <Name>-san.
Names are resolved against data/members.json so IPs and workarea prefixes
("FI: ", "CO: ", "OP: ") match what index.html expects.
"""
import argparse, json, re, sys
from pathlib import Path
import openpyxl

ROOT = Path(__file__).resolve().parent.parent
DAYS = ['MON', 'TUE', 'WED', 'THU', 'FRI']
DAY_FULL = dict(zip(DAYS, ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']))
LABELS = {  # team -> workarea prefix (mirrors HIERARCHY in index.html)
    **{t: 'FI' for t in ['ODMS', 'OTC Derivative', 'IR Server', 'Research', 'FIQA']},
    **{t: 'CO' for t in ['Cash Management', 'Accounting', 'Master', 'MarketDB']},
    **{t: 'OP' for t in ['Risk IT', 'Risk EUC']},
}
# sheet team label -> members.json team
TEAM_ALIAS = {
    'riskit': 'Risk IT', 'risk euc': 'Risk EUC', 'irservr': 'IR Server', 'pega': 'PEGA',
    'marketdb': 'MarketDB', 'accounting': 'Accounting', 'odms': 'ODMS', 'qa': 'FIQA',
    'master': 'Master', 'cash mngmt': 'Cash Management', 'deriva': 'OTC Derivative',
    'research': 'Research', 'karen-san': 'MarketDB', 'muro-san': 'ODMS', 'imee-san': 'PMO',
}
# sheet nickname -> members.json nickname (only where they differ)
NICK_ALIAS = {
    'patrick p': 'Patrick', 'neilb': 'Neil', 'neiljeff': 'Jeff', 'jay-ar': 'Jay-ar',
    'stingray': 'Sting', 'aubrey': 'Abby', 'patrick g': 'PatrickG', 'patrick g': 'PatrickG',
}


def resolve(team_raw, nick_raw, roster):
    team = TEAM_ALIAS.get(team_raw.strip().lower())
    nick = NICK_ALIAS.get(nick_raw.strip().lower(), nick_raw.strip())
    if team == 'PEGA' and not nick.startswith('PEGA-'):
        nick = 'PEGA-' + nick
    cands = [m for m in roster if m['nickname'].lower() == nick.lower()]
    if len(cands) > 1 and team:
        cands = [m for m in cands if m['team'] == team] or cands
    return cands[0] if cands else None


def cell_to_entry(val, roster, warnings, where, added, notes_ips):
    if val is None or str(val).strip() in ('', 'FREE', 'NA'):
        return {'name': '', 'ip': '', 'unavailable': False}
    text = str(val).strip()
    if text.upper().startswith('RESERVED:'):  # sheet counts held desks as unoccupied
        return {'name': '', 'ip': '', 'unavailable': False}
    parts = text.split('\n')
    team_raw, nick_raw = (parts[0], parts[-1]) if len(parts) > 1 else (parts[0], parts[0])
    m = resolve(team_raw, nick_raw, roster)
    team = TEAM_ALIAS.get(team_raw.strip().lower())
    if not m and team and not re.match(r'TBD\d*$', nick_raw.strip()):
        nick = NICK_ALIAS.get(nick_raw.strip().lower(), nick_raw.strip())
        m = {'name': nick, 'nickname': nick, 'team': team, 'ip': notes_ips.get(nick.lower(), '')}
        roster.append(m)
        added.append(m)
    if not m:
        warnings.append(f'{where}: "{text!r}" not in members.json - stored as plain text')
        return {'name': nick_raw.strip(), 'ip': '', 'unavailable': False}
    label = LABELS.get(m['team'])
    return {'name': f"{label}: {m['nickname']}" if label else m['nickname'],
            'ip': m['ip'], 'unavailable': False}


def notes_table(ws):
    """first-name -> IP from the 'Added / Not yet added' table (cols E:G, below the 'Seat' header)."""
    ips, in_table = {}, False
    for r in range(1, ws.max_row + 1):
        e = str(ws.cell(r, 5).value or '').strip()
        if e == 'Seat':
            in_table = True
        if not in_table or not re.fullmatch(r'\d\d\.\d\d[A-H]', e):
            continue
        names = str(ws.cell(r, 6).value or '').split(chr(10))
        addrs = str(ws.cell(r, 7).value or '').split(chr(10))
        for n, ip in zip(names, addrs):
            if n.strip() and ip.strip():
                ips.setdefault(n.strip().lower(), ip.strip())
    return ips


def parse_sheet(ws, roster):
    seats, warnings, added = {}, [], []
    notes_ips = notes_table(ws)
    na_count = 0
    for r in range(1, ws.max_row + 1):
        # a seat header row has a day label in col B on the NEXT row and a seat id in D..N
        if str(ws.cell(r + 1, 2).value or '').strip().upper() != 'MON':
            continue
        for c in range(3, ws.max_column + 1):
            sid = ws.cell(r, c).value
            if not sid or not re.fullmatch(r'\d\d\.\d\d[A-H]|NA', str(sid).strip()):
                continue
            sid = str(sid).strip()
            if sid == 'NA':
                na_count += 1
                sid = 'NA2' if na_count == 1 else 'NA'
                seats[sid] = {DAY_FULL[d]: {'name': '', 'ip': '', 'unavailable': True} for d in DAYS}
                continue
            seats[sid] = {}
            for i, d in enumerate(DAYS, 1):
                seats[sid][DAY_FULL[d]] = cell_to_entry(
                    ws.cell(r + i, c).value, roster, warnings, f'{sid} {d}', added, notes_ips)
    return seats, warnings, added


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('xlsx')
    ap.add_argument('--sheet', default='Seat Layout - Latest')
    ap.add_argument('--dry-run', action='store_true')
    a = ap.parse_args()

    roster = json.loads((ROOT / 'data/members.json').read_text('utf-8'))['members']
    ws = openpyxl.load_workbook(a.xlsx, data_only=True)[a.sheet]
    seats, warnings, added = parse_sheet(ws, roster)

    html = (ROOT / 'index.html').read_text('utf-8')
    expected = re.search(r'const ALL_SEATS=\[(.*?)\];', html, re.S).group(1)
    expected = re.findall(r"'([^']+)'", expected)
    missing = [s for s in expected if s not in seats]
    extra = [s for s in seats if s not in expected]
    if missing or extra:
        sys.exit(f'Seat mismatch with index.html ALL_SEATS - missing {missing}, unexpected {extra}')

    for m in added:
        print('ADDED member', m['nickname'], m['team'], m['ip'] or '(no IP in sheet)')
    for w in warnings:
        print('WARN', w)
    free = {d: sum(1 for s in seats.values() if not s[DAY_FULL[d]]['name'] and not s[DAY_FULL[d]]['unavailable'])
            for d in DAYS}
    print('Unoccupied seats per day:', free)
    if a.dry_run:
        return
    if added:
        mp = ROOT / 'data/members.json'
        mp.write_text(json.dumps({'members': roster}, indent=2) + '\n', 'utf-8')
        print('Wrote data/members.json')
    out = {'seats': {s: seats[s] for s in expected}}
    (ROOT / 'data/schedule.json').write_text(json.dumps(out, indent=2) + '\n', 'utf-8')
    print('Wrote data/schedule.json')


if __name__ == '__main__':
    main()
