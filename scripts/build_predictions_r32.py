#!/usr/bin/env python3
"""Build RondaEliminatoria/data/predictions.json from the Excel.

Parses the "Ronda Eliminatoria" block (16 Round-of-32 matches) of
`Porra Mundial 2026.xlsx`. Each player cell is an exact-score prediction like
`4-1`, optionally annotated with who advances on a draw (`2-2 (ALE)`,
`1-1 CAN`, `2-2 (pen Brasil)`). Emits players, the 16 matches in bracket order
(with Spanish display names + team codes), and per-player {g1, g2, adv}.
"""
import json
import re
import sys
import unicodedata
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "RondaEliminatoria" / "Porra Mundial 2026.xlsx"
OUT = ROOT / "RondaEliminatoria" / "data" / "predictions.json"

# Bracket order is the row order in the Excel. side/pos drive the layout:
# 8 matches down the left, 8 down the right, top -> bottom.
SIDES = (["left"] * 8) + (["right"] * 8)

# Spanish display name -> team code (codes match ../data/team-codes.json).
NAME2CODE = {
    "alemania": "GER", "paraguay": "PAR", "francia": "FRA", "suecia": "SWE",
    "sudafrica": "RSA", "canada": "CAN", "paisesbajos": "NED",
    "marruecos": "MAR", "portugal": "POR", "croacia": "CRO", "espana": "ESP",
    "austria": "AUT", "eeuu": "USA", "bosnia": "BOS", "belgica": "BEL",
    "senegal": "SEN", "brasil": "BRA", "japon": "JPN", "cdemarfil": "CIV",
    "cmarfil": "CIV", "noruega": "NOR", "mexico": "MEX", "ecuador": "ECU",
    "inglaterra": "ENG", "rdcongo": "CON", "argentina": "ARG",
    "caboverde": "CPV", "australia": "AUS", "egipto": "EGY", "suiza": "SUI",
    "argelia": "ALG", "colombia": "COL", "ghana": "GHA",
}


def norm(s):
    """Lowercase, strip accents and non-letters — for fuzzy name matching."""
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z]", "", s.lower())


def name_to_code(name):
    code = NAME2CODE.get(norm(name))
    if not code:
        sys.exit(f"Unknown team name: {name!r} (normalized {norm(name)!r})")
    return code


def resolve_advancer(token, t1, t2, es1, es2):
    """Pick which of the two teams an annotation token refers to.

    Scoped to the two teams in the row, so a prefix/code match is unambiguous.
    Matches against the team code and the normalized Spanish name.
    """
    tok = norm(token.replace("pen", ""))
    if not tok:
        return None
    for code, es in ((t1, es1), (t2, es2)):
        n = norm(es)
        if tok == code.lower() or n.startswith(tok) or tok.startswith(code.lower()):
            return code
    return None


def parse_cell(raw, t1, t2, es1, es2, ctx):
    """Return {g1, g2, adv} or None for a blank cell."""
    if raw is None or str(raw).strip() == "":
        return None
    s = str(raw).strip()
    m = re.search(r"(\d+)\s*-\s*(\d+)", s)
    if not m:
        print(f"  WARN {ctx}: cannot parse score from {s!r}", file=sys.stderr)
        return None
    g1, g2 = int(m.group(1)), int(m.group(2))
    annotation = (s[: m.start()] + s[m.end():]).strip()
    if g1 > g2:
        adv = t1
    elif g2 > g1:
        adv = t2
    else:
        adv = resolve_advancer(annotation, t1, t2, es1, es2)
        if adv is None:
            print(f"  WARN {ctx}: drawn pick {s!r} has no advancer", file=sys.stderr)
    return {"g1": g1, "g2": g2, "adv": adv}


def main():
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ws = wb["Sheet1"]

    # Locate the "Ronda Eliminatoria" header row, then read the 16 match rows.
    header_row = None
    for r in range(1, ws.max_row + 1):
        if norm(ws.cell(r, 1).value or "") == "rondaeliminatoria":
            header_row = r
            break
    if header_row is None:
        sys.exit("Could not find 'Ronda Eliminatoria' section in the Excel")

    # Players come from row 1 (skip the "Partidos" corner cell).
    players = []
    c = 2
    while ws.cell(1, c).value not in (None, ""):
        players.append(str(ws.cell(1, c).value).strip())
        c += 1

    matches, predictions, names = [], {}, {}
    row = header_row + 1
    idx = 0
    while idx < 16:
        label = ws.cell(row, 1).value
        if label is None or str(label).strip() == "":
            row += 1
            continue
        es1, es2 = [p.strip() for p in str(label).split(" - ", 1)]
        t1, t2 = name_to_code(es1), name_to_code(es2)
        mid = f"r32-{idx + 1:02d}"
        matches.append({"id": mid, "side": SIDES[idx], "pos": idx % 8,
                        "team1": t1, "team2": t2})
        names[t1], names[t2] = es1, es2
        predictions[mid] = {}
        for i, player in enumerate(players):
            cell = ws.cell(row, 2 + i).value
            pick = parse_cell(cell, t1, t2, es1, es2, f"{mid} {player}")
            if pick is not None:
                predictions[mid][player] = pick
        idx += 1
        row += 1

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(
        {"players": players, "matches": matches, "names": names,
         "predictions": predictions},
        ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Wrote {OUT.relative_to(ROOT)}: {len(players)} players, "
          f"{len(matches)} matches")


if __name__ == "__main__":
    main()
