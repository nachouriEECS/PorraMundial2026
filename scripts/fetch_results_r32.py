#!/usr/bin/env python3
"""Fetch real Round-of-32 results and write RondaEliminatoria/data/results.json.

Mirrors scripts/fetch_results.py but for the knockout bracket, where we need the
exact full-time goals and who ADVANCES (penalty winner on a draw), not just 1/X/2.

Source: football-data.org v4, competition WC. Token in FOOTBALL_DATA_TOKEN.
Each porra fixture is matched to a real knockout fixture by the unordered pair of
team codes (group-stage fixtures are ignored so a group rematch can't collide).

Usage:
  FOOTBALL_DATA_TOKEN=xxx python3 scripts/fetch_results_r32.py            # write
  FOOTBALL_DATA_TOKEN=xxx python3 scripts/fetch_results_r32.py --verify   # no write
"""
import http.client
import json
import os
import sys
import time
import unicodedata
import urllib.request
import urllib.error
from datetime import datetime, timezone
from pathlib import Path

API_URL = "https://api.football-data.org/v4/competitions/WC/matches"
MAX_ATTEMPTS = 4
RETRY_BACKOFF = 5

ROOT = Path(__file__).resolve().parent.parent
RDIR = ROOT / "RondaEliminatoria" / "data"
PRED = RDIR / "predictions.json"
CODES = ROOT / "data" / "team-codes.json"
OVERRIDES = RDIR / "overrides.json"
OUT = RDIR / "results.json"


def norm(s):
    if not s:
        return ""
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = "".join(c if c.isalnum() or c.isspace() else " " for c in s.lower())
    return " ".join(s.split())


def build_alias_index(codes):
    idx = {}
    for code, info in codes.items():
        for n in [info["name"]] + info.get("aliases", []):
            idx[norm(n)] = code
    return idx


def resolve_code(name, alias_idx):
    n = norm(name)
    if n in alias_idx:
        return alias_idx[n]
    for alias, code in alias_idx.items():
        if alias and (alias in n or n in alias):
            return code
    return None


def fetch_api(token):
    req = urllib.request.Request(API_URL, headers={"X-Auth-Token": token})
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", "replace")
            if e.code != 429 and e.code < 500:
                sys.exit(f"API HTTP {e.code}: {body}")
            err = f"API HTTP {e.code}: {body}"
        except (urllib.error.URLError, http.client.HTTPException, OSError) as e:
            err = f"API request failed: {e}"
        if attempt < MAX_ATTEMPTS:
            wait = RETRY_BACKOFF * attempt
            print(f"{err}\n  retrying in {wait}s ({attempt}/{MAX_ATTEMPTS - 1})...", file=sys.stderr)
            time.sleep(wait)
    sys.exit(f"{err}\n  giving up after {MAX_ATTEMPTS} attempts.")


def main():
    verify = "--verify" in sys.argv
    token = os.environ.get("FOOTBALL_DATA_TOKEN", "").strip()
    if not token:
        sys.exit("FOOTBALL_DATA_TOKEN is not set.")

    pred = json.loads(PRED.read_text(encoding="utf-8"))
    codes = json.loads(CODES.read_text(encoding="utf-8"))
    overrides = json.loads(OVERRIDES.read_text(encoding="utf-8")).get("results", {})
    alias_idx = build_alias_index(codes)

    api_matches = fetch_api(token).get("matches", [])

    # Index knockout fixtures by the unordered pair of resolved codes. Group-stage
    # matches are skipped so a knockout rematch of a group pairing can't collide.
    by_pair, unresolved = {}, set()
    for m in api_matches:
        if (m.get("stage") or "").upper() == "GROUP_STAGE":
            continue
        h = m.get("homeTeam", {}).get("name")
        a = m.get("awayTeam", {}).get("name")
        hc, ac = resolve_code(h, alias_idx), resolve_code(a, alias_idx)
        if not hc:
            unresolved.add(h)
        if not ac:
            unresolved.add(a)
        if not hc or not ac:
            continue
        score = m.get("score") or {}
        ft = score.get("fullTime") or {}
        pen = score.get("penalties") or {}
        by_pair[frozenset((hc, ac))] = {
            "home": hc, "away": ac, "status": m.get("status"),
            "winner": score.get("winner"), "utcDate": m.get("utcDate"),
            "hg": ft.get("home"), "ag": ft.get("away"),
            "ph": pen.get("home"), "pa": pen.get("away"),
        }

    results, unmatched, finished = {}, [], []
    for match in pred["matches"]:
        mid, t1, t2 = match["id"], match["team1"], match["team2"]
        entry = {"status": None, "g1": None, "g2": None, "adv": None,
                 "score": None, "utcDate": None, "source": "api"}
        fx = by_pair.get(frozenset((t1, t2)))
        if fx:
            entry["status"], entry["utcDate"] = fx["status"], fx["utcDate"]
            flip = fx["home"] != t1   # orient API home/away to porra team1/team2
            g1, g2 = (fx["ag"], fx["hg"]) if flip else (fx["hg"], fx["ag"])
            ph, pa = (fx["pa"], fx["ph"]) if flip else (fx["ph"], fx["pa"])
            if fx["status"] == "FINISHED" and g1 is not None and g2 is not None:
                entry["g1"], entry["g2"] = g1, g2
                entry["score"] = f"{g1}-{g2}"
                if g1 > g2:
                    entry["adv"] = t1
                elif g2 > g1:
                    entry["adv"] = t2
                elif ph is not None and pa is not None and ph != pa:
                    entry["adv"] = t1 if ph > pa else t2
                elif fx["winner"] in ("HOME_TEAM", "AWAY_TEAM"):
                    home_is_t1 = not flip
                    entry["adv"] = t1 if (fx["winner"] == "HOME_TEAM") == home_is_t1 else t2
        else:
            unmatched.append(mid)

        # Manual override wins over the API. Expected: {g1, g2, adv}.
        ov = overrides.get(mid)
        if isinstance(ov, dict) and ov.get("g1") is not None and ov.get("g2") is not None:
            entry.update(status="FINISHED", g1=ov["g1"], g2=ov["g2"],
                         adv=ov.get("adv"), score=f"{ov['g1']}-{ov['g2']}", source="override")

        if entry["g1"] is not None:
            finished.append((entry["utcDate"] or "", mid))
        results[mid] = entry

    finished.sort()
    last_finished = finished[-1][1] if finished else None
    upcoming = sorted((e["utcDate"], mid) for mid, e in results.items()
                      if e["g1"] is None and e["status"] != "FINISHED" and e["utcDate"])
    next_match = upcoming[0][1] if upcoming else None

    out = {
        "lastUpdated": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "lastFinishedId": last_finished,
        "nextMatchId": next_match,
        "matches": results,
    }

    print(f"API matches fetched: {len(api_matches)}")
    print(f"Porra matches resolved to a fixture: {len(pred['matches']) - len(unmatched)}/{len(pred['matches'])}")
    print(f"Matches with a final result: {len(finished)}")
    if unresolved:
        print("\n!! API team names that did NOT map to a code:")
        for n in sorted(unresolved):
            print(f"   - {n!r}")
    if unmatched:
        print("\n!! Porra matches with no knockout fixture yet (use overrides.json if needed):")
        print("   " + ", ".join(unmatched))

    if verify:
        print("\n--verify: not writing results.json")
        return
    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"\nWrote {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
