#!/usr/bin/env python3
"""Build data/predictions.json from the porra Excel sheet.

Run occasionally (only when the predictions change). The .xlsx is just a zip of
XML, so we read it with the standard library — no pandas/openpyxl needed.

Layout of the sheet (Sheet1):
  - Row 1: A = "Partidos", columns B..P = player names (15 players)
  - Match rows: column A = "MEX - RSA" (team1 code - team2 code),
    columns B..P = each player's pick: 1.0 (team1), 2.0 (team2) or "X" (draw)
  - Blocks are separated by "Jornada 1/2/3" header rows; "PUNTOS" ends the matches.

Output: data/predictions.json
  {
    "players": ["Javi", ...],
    "matches": [{"id": "j1-01", "matchday": 1, "team1": "MEX", "team2": "RSA"}, ...],
    "predictions": {"j1-01": {"Javi": "2", "Nacho": "1", ...}, ...}
  }
Picks are normalised to the strings "1" | "X" | "2"; a missing pick is null.
"""
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT.parent / "Porra Mundial 2026.xlsx"
OUT = ROOT / "data" / "predictions.json"


def col_to_num(col):
    """'A' -> 1, 'B' -> 2, ... 'AA' -> 27."""
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def read_sheet(xlsx_path):
    """Return {row_number: {col_letter: value}} for Sheet1."""
    z = zipfile.ZipFile(xlsx_path)
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root:
            shared.append("".join(t.text or "" for t in si.iter(NS + "t")))

    sheet = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
    rows = {}
    for c in sheet.iter(NS + "c"):
        ref = c.attrib["r"]
        m = re.match(r"([A-Z]+)(\d+)", ref)
        col, row = m.group(1), int(m.group(2))
        t = c.attrib.get("t")
        v = c.find(NS + "v")
        val = None
        if t == "inlineStr":
            is_ = c.find(NS + "is")
            if is_ is not None:
                val = "".join(tt.text or "" for tt in is_.iter(NS + "t"))
        elif v is not None:
            val = v.text
            if t == "s":
                val = shared[int(val)]
        rows.setdefault(row, {})[col] = val
    return rows


def normalise_pick(raw):
    """1.0/1 -> '1', 2.0/2 -> '2', X/x -> 'X', else None."""
    if raw is None:
        return None
    s = str(raw).strip()
    if s == "":
        return None
    if s.upper() == "X":
        return "X"
    try:
        n = int(float(s))
    except ValueError:
        return None
    if n == 1:
        return "1"
    if n == 2:
        return "2"
    return None


def main():
    if not XLSX.exists():
        sys.exit(f"Excel file not found: {XLSX}")

    rows = read_sheet(XLSX)

    # Players: header row 1, columns B onwards until blank.
    header = rows[1]
    player_cols = []  # [(col_letter, name)]
    col = "B"
    while True:
        name = header.get(col)
        if not name:
            break
        player_cols.append((col, str(name).strip()))
        col = num_to_col(col_to_num(col) + 1)
    players = [name for _, name in player_cols]

    matches = []
    predictions = {}
    matchday = 0
    idx_in_day = 0

    for row in sorted(rows):
        if row == 1:
            continue
        a = rows[row].get("A")
        if not a:
            continue
        a = str(a).strip()

        jm = re.match(r"Jornada\s+(\d+)", a, re.IGNORECASE)
        if jm:
            matchday = int(jm.group(1))
            idx_in_day = 0
            continue

        # Stop once we reach the bonus section.
        if a.upper() in ("PUNTOS", "GANADORES", "PREGUNTAS"):
            break

        mm = re.match(r"^([A-Z]{2,4})\s*-\s*([A-Z]{2,4})$", a)
        if not mm:
            continue  # not a match row

        idx_in_day += 1
        match_id = f"j{matchday}-{idx_in_day:02d}"
        team1, team2 = mm.group(1), mm.group(2)
        matches.append(
            {"id": match_id, "matchday": matchday, "team1": team1, "team2": team2}
        )
        picks = {}
        for col, name in player_cols:
            picks[name] = normalise_pick(rows[row].get(col))
        predictions[match_id] = picks

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(
        json.dumps(
            {"players": players, "matches": matches, "predictions": predictions},
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )

    print(f"Wrote {OUT}")
    print(f"  players: {len(players)}")
    print(f"  matches: {len(matches)}")
    by_day = {}
    for m in matches:
        by_day[m["matchday"]] = by_day.get(m["matchday"], 0) + 1
    print(f"  per matchday: {dict(sorted(by_day.items()))}")


def num_to_col(n):
    """1 -> 'A', 27 -> 'AA'."""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(ord("A") + r) + s
    return s


if __name__ == "__main__":
    main()
