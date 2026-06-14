# Porra Mundial 2026 — live leaderboard

A self-updating leaderboard for our World Cup 2026 group-stage sweepstake. Everyone's
`1` / `X` / `2` predictions are scored **1 point per correct group-stage result**. A
scheduled GitHub Action pulls the real results and the page refreshes itself.

- **`index.html`** — the leaderboard (loads the JSON files, computes points in the browser).
- **`data/predictions.json`** — everyone's picks, generated from the Excel.
- **`data/team-codes.json`** — maps the 3-letter codes (`MEX`, `RSA`, …) to country names + flags.
- **`data/results.json`** — real results, written by the Action. **Don't edit by hand.**
- **`data/overrides.json`** — manual safety net (see below).
- **`scripts/build_predictions.py`** — rebuilds `predictions.json` from `Porra Mundial 2026.xlsx`.
- **`scripts/fetch_results.py`** — fetches results from football-data.org.
- **`.github/workflows/update-results.yml`** — runs the fetch every 30 min.

## One-time setup

### 1. Get a free football-data.org API token
Register at <https://www.football-data.org/client/register>. The **free tier includes the
World Cup** (10 requests/min). You'll get an API token by email.

### 2. Create the repo (must be **public** for free Pages + Actions)
```bash
cd porra-mundial-2026
git init
git add .
git commit -m "Porra Mundial 2026 leaderboard"
gh repo create porra-mundial-2026 --public --source=. --push
# or create the repo on github.com and: git remote add origin … && git push -u origin main
```

### 3. Add the API token as a secret
Repo → **Settings → Secrets and variables → Actions → New repository secret**:
- **Name:** `FOOTBALL_DATA_TOKEN`
- **Value:** *(the token from step 1)*

### 4. Enable GitHub Pages
Repo → **Settings → Pages** → *Build and deployment* → **Deploy from a branch** →
branch **`main`**, folder **`/ (root)`** → **Save**.

Your link (share this with everyone): **`https://<your-username>.github.io/porra-mundial-2026/`**

### 5. Run the Action once
Repo → **Actions** → *Update results* → **Run workflow**. It will fetch results and commit
`data/results.json`. After that it runs automatically every 30 minutes.

## Updating predictions
If a pick changes in the Excel, rebuild and push:
```bash
python3 scripts/build_predictions.py
git commit -am "Update predictions" && git push
```

## Manual override (safety net)
If a result is wrong, missing, or the API can't match a fixture, set it by hand in
`data/overrides.json`:
```json
{ "results": { "j1-03": "1", "j2-10": "X" } }
```
Keys are match ids (`j<matchday>-<NN>`, in fixture order); values are `"1"` (team 1 wins),
`"X"` (draw) or `"2"` (team 2 wins). Overrides always win over the API. Commit and push, or
just re-run the Action.

## Test the fetch locally
```bash
export FOOTBALL_DATA_TOKEN=your_token_here
python3 scripts/fetch_results.py --verify   # diagnostics, writes nothing
python3 scripts/fetch_results.py            # actually writes data/results.json
```
`--verify` reports how many Excel matches resolved to a real fixture and lists any team names
that didn't map (add them to `team-codes.json` aliases) — useful the first time to confirm the
mapping is complete.
