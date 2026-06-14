#!/usr/bin/env python3
"""Fetch real World Cup results and write data/results.json.

Called by the GitHub Action on a schedule. Pure standard library so the Action
needs no `pip install`.

Source: football-data.org v4, competition WC. Set the API token in the
environment variable FOOTBALL_DATA_TOKEN (a GitHub Actions secret).

Resolution strategy: each World Cup group pairing is unique, so we match each
Excel match to its real fixture by the *unordered* pair of team codes, then
orient the outcome to the Excel team1/team2. The result for each match is one of
"1" (team1 won), "X" (draw), "2" (team2 won), or null (not finished yet).

Usage:
  FOOTBALL_DATA_TOKEN=xxx python3 scripts/fetch_results.py            # write results.json
  FOOTBALL_DATA_TOKEN=xxx python3 scripts/fetch_results.py --verify   # diagnostics, no write
"""
import json
import os
import sys
import unicodedata
import urllib.request
import urllib.error
from datetime import datetime, timezone
from pathlib import Path

API_URL = "https://api.football-data.org/v4/competitions/WC/matches"

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
PRED = DATA / "predictions.json"
CODES = DATA / "team-codes.json"
OVERRIDES = DATA / "overrides.json"
OUT = DATA / "results.json"


def norm(s):
    """Lowercase, strip accents and punctuation, collapse spaces."""
    if not s:
        return ""
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower()
    s = "".join(c if c.isalnum() or c.isspace() else " " for c in s)
    return " ".join(s.split())


def build_alias_index(codes):
    """Map normalised team name/alias -> 3-letter code."""
    idx = {}
    for code, info in codes.items():
        names = [info["name"]] + info.get("aliases", [])
        for n in names:
            idx[norm(n)] = code
    return idx


def fetch_api(token):
    req = urllib.request.Request(API_URL, headers={"X-Auth-Token": token})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", "replace")
        sys.exit(f"API HTTP {e.code}: {body}")
    except urllib.error.URLError as e:
        sys.exit(f"API request failed: {e}")


def resolve_code(name, alias_idx):
    """Return the team code for an API team name, or None."""
    n = norm(name)
    if n in alias_idx:
        return alias_idx[n]
    # Fallback: substring / token containment for stubborn names.
    for alias, code in alias_idx.items():
        if alias and (alias in n or n in alias):
            return code
    return None


def main():
    verify = "--verify" in sys.argv
    token = os.environ.get("FOOTBALL_DATA_TOKEN", "").strip()
    if not token:
        sys.exit("FOOTBALL_DATA_TOKEN is not set.")

    pred = json.loads(PRED.read_text(encoding="utf-8"))
    codes = json.loads(CODES.read_text(encoding="utf-8"))
    overrides = json.loads(OVERRIDES.read_text(encoding="utf-8")).get("results", {})
    alias_idx = build_alias_index(codes)

    payload = fetch_api(token)
    api_matches = payload.get("matches", [])

    # Index API fixtures by the unordered pair of resolved codes.
    by_pair = {}
    unresolved_names = set()
    for m in api_matches:
        h = m.get("homeTeam", {}).get("name")
        a = m.get("awayTeam", {}).get("name")
        hc = resolve_code(h, alias_idx)
        ac = resolve_code(a, alias_idx)
        if not hc:
            unresolved_names.add(h)
        if not ac:
            unresolved_names.add(a)
        if not hc or not ac:
            continue
        by_pair[frozenset((hc, ac))] = {
            "home": hc, "away": ac,
            "status": m.get("status"),
            "winner": (m.get("score") or {}).get("winner"),
            "homeGoals": ((m.get("score") or {}).get("fullTime") or {}).get("home"),
            "awayGoals": ((m.get("score") or {}).get("fullTime") or {}).get("away"),
            "utcDate": m.get("utcDate"),
        }

    results = {}
    unmatched = []
    finished = []
    for match in pred["matches"]:
        mid, t1, t2 = match["id"], match["team1"], match["team2"]
        entry = {"status": None, "outcome": None, "score": None, "utcDate": None, "source": "api"}
        fx = by_pair.get(frozenset((t1, t2)))
        if fx:
            entry["status"] = fx["status"]
            entry["utcDate"] = fx["utcDate"]
            # Orient goals/winner to the Excel team1 vs team2.
            if fx["home"] == t1:
                g1, g2, win_is_t1 = fx["homeGoals"], fx["awayGoals"], "HOME_TEAM"
            else:
                g1, g2, win_is_t1 = fx["awayGoals"], fx["homeGoals"], "AWAY_TEAM"
            if fx["status"] == "FINISHED":
                # Prefer the goals (source of truth); fall back to the winner field.
                # The free tier sometimes marks a match FINISHED with a null winner.
                if g1 is not None and g2 is not None:
                    entry["outcome"] = "1" if g1 > g2 else "2" if g2 > g1 else "X"
                    entry["score"] = f"{g1}-{g2}"
                elif fx["winner"]:
                    entry["outcome"] = ("X" if fx["winner"] == "DRAW"
                                        else "1" if fx["winner"] == win_is_t1 else "2")
        else:
            unmatched.append(mid)

        # Manual override wins over the API.
        if mid in overrides and overrides[mid] in ("1", "X", "2"):
            entry["outcome"] = overrides[mid]
            entry["status"] = "FINISHED"
            entry["source"] = "override"

        if entry["outcome"] is not None:
            finished.append((entry["utcDate"] or "", mid))
        results[mid] = entry

    # Most recently finished match (by kickoff time), for the header highlight.
    last_finished = None
    if finished:
        finished.sort()
        last_finished = finished[-1][1]

    # Next match: earliest not-yet-decided fixture by kickoff time.
    upcoming = sorted(
        (e["utcDate"], mid) for mid, e in results.items()
        if e["outcome"] is None and e["status"] != "FINISHED" and e["utcDate"]
    )
    next_match = upcoming[0][1] if upcoming else None

    out = {
        "lastUpdated": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "lastFinishedId": last_finished,
        "nextMatchId": next_match,
        "matches": results,
    }

    print(f"API matches fetched: {len(api_matches)}")
    print(f"Excel matches resolved to a fixture: {len(pred['matches']) - len(unmatched)}/{len(pred['matches'])}")
    print(f"Matches with a final result: {len(finished)}")
    if unresolved_names:
        print("\n!! API team names that did NOT map to a code (add to team-codes.json aliases):")
        for n in sorted(unresolved_names):
            print(f"   - {n!r}")
    if unmatched:
        print("\n!! Excel matches with no fixture found (use overrides.json if needed):")
        print("   " + ", ".join(unmatched))

    if verify:
        print("\n--verify: not writing results.json")
        return

    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"\nWrote {OUT}")


if __name__ == "__main__":
    main()
