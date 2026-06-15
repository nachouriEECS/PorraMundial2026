#!/usr/bin/env python3
"""Build data/predictions.json from the porra Excel sheet.

Run occasionally (only when the predictions change). The .xlsx is just a zip of
XML, so we read it with the standard library — no pandas/openpyxl needed.

Layout of the sheet (Sheet1):
  - Row 1: A = "Partidos", columns B..P = player names (15 players)
  - Match rows: column A = "MEX - RSA" (team1 code - team2 code),
    columns B..P = each player's pick: 1.0 (team1), 2.0 (team2) or "X" (draw)
  - Blocks are separated by "Jornada 1/2/3" header rows; "PUNTOS" ends the matches.
  - Below the matches, a bonus section has one row per question (column A = the
    question label, e.g. "GANADORES", "Pichichi"); columns B..P = each player's
    answer.

Output: data/predictions.json
  {
    "players": ["Javi", ...],
    "matches": [{"id": "j1-01", "matchday": 1, "team1": "MEX", "team2": "RSA"}, ...],
    "predictions": {"j1-01": {"Javi": "2", "Nacho": "1", ...}, ...},
    "bonusQuestions": [{"key": "ganadores", "label": "Campeón"}, ...],
    "bonus": {"Javi": {"ganadores": "Holanda", ...}, ...}
  }
Picks are normalised to the strings "1" | "X" | "2"; a missing pick is null.
A missing bonus answer is null.
"""
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

ROOT = Path(__file__).resolve().parent.parent
# The workbook may live in the project root or one level up; use whichever exists.
XLSX_NAME = "Porra Mundial 2026.xlsx"
XLSX = next(
    (p for p in (ROOT / XLSX_NAME, ROOT.parent / XLSX_NAME) if p.exists()),
    ROOT / XLSX_NAME,
)
OUT = ROOT / "data" / "predictions.json"

# Bonus questions, in display order. Each entry maps an output key + friendly
# label to the column-A text used in the sheet (matched case-insensitively).
# `numeric` answers are normalised to a plain integer string ("9.0" -> "9").
BONUS_QUESTIONS = [
    {"key": "ganadores", "label": "Campeón", "match": "ganadores", "numeric": False},
    {"key": "mejorGol", "label": "Mejor gol", "match": "mejor gol", "numeric": False},
    {"key": "pichichi", "label": "Pichichi", "match": "pichichi", "numeric": False},
    {"key": "golesPropia", "label": "Goles en propia", "match": "goles en propia", "numeric": True},
    {"key": "masGoleada", "label": "Más goleada", "match": "mas goleada", "numeric": False},
    {"key": "banda", "label": "Nº de banda", "match": "# banda", "numeric": True},
]


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


def normalise_bonus(raw, numeric):
    """Clean a bonus answer; numeric answers become a plain int string.

    "9.0" -> "9" when numeric; text answers are just stripped; blanks -> None.
    """
    if raw is None:
        return None
    s = str(raw).strip()
    if s == "":
        return None
    if numeric:
        try:
            return str(int(float(s)))
        except ValueError:
            return s
    return s


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
    bonus = {name: {} for name in players}
    matchday = 0
    idx_in_day = 0
    in_bonus = False  # flipped once we pass "PUNTOS" into the bonus section

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

        # "PUNTOS" marks the end of the matches and the start of the bonus block.
        if a.upper() == "PUNTOS":
            in_bonus = True
            continue

        if in_bonus:
            q = next((q for q in BONUS_QUESTIONS if q["match"] == a.lower()), None)
            if q:
                for col, name in player_cols:
                    bonus[name][q["key"]] = normalise_bonus(
                        rows[row].get(col), q["numeric"]
                    )
            continue

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

    # Ensure every player has an entry for every question (missing -> None).
    for name in players:
        for q in BONUS_QUESTIONS:
            bonus[name].setdefault(q["key"], None)

    bonus_questions = [{"key": q["key"], "label": q["label"]} for q in BONUS_QUESTIONS]

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(
        json.dumps(
            {
                "players": players,
                "matches": matches,
                "predictions": predictions,
                "bonusQuestions": bonus_questions,
                "bonus": bonus,
            },
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
