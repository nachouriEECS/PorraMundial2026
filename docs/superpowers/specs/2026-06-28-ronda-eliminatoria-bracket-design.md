# Ronda Eliminatoria (Round of 32) bracket site — design

## Goal

A second self-updating page for the World Cup 2026 sweepstake, covering the
**Round of 32** (16 knockout matches). Unlike the group stage (1/X/2 picks,
1 pt each), here every player predicts an **exact score**, scored **0–3 pts per
match**. The page shows each player's predictions laid out on a bracket that
mirrors the official Round-of-32 graphic, with the selected player's name and
total points in the center, and per-match boxes color-coded by points earned.

Lives at `RondaEliminatoria/index.html` →
`https://nachourieecs.github.io/PorraMundial2026/RondaEliminatoria/`.
The existing group-stage site at the repo root is untouched.

## Scoring (0–3 per match, computed in the browser)

For each match and player:

- **+1** if the predicted **advancer** matches the real advancer. In the
  knockout round the advancer includes penalty-shootout resolution: a predicted
  `2-2 (ALE)` earns this point if Germany advances.
- **+1** if predicted **team-1 goals** = real team-1 goals.
- **+1** if predicted **team-2 goals** = real team-2 goals.

Goals are compared against the **full-time score** (after extra time if any,
before the penalty shootout). Max 3 per match.

Worked example: real `4-1`, predicted `3-1` → advancer ✓ (+1), team-1 `3≠4`
(0), team-2 `1=1` (+1) = **2 pts**.

A player's total = sum over the 16 matches. The leaderboard ranks by total,
tie-broken by name (Spanish locale).

## Data model

New `RondaEliminatoria/data/` folder. Reuses the root `../data/team-codes.json`
for flags + names.

### `predictions.json` (built from the Excel)
```jsonc
{
  "players": ["Javi", "Nacho", ...],            // 15
  "matches": [
    { "id": "r32-01", "side": "left",  "pos": 0,
      "team1": "GER", "team2": "PAR" },          // 16 matches, bracket order
    ...
  ],
  "predictions": {
    "r32-01": {
      "Javi": { "g1": 1, "g2": 0, "adv": "GER" },
      ...
    }
  }
}
```
- `side`/`pos` place the match in the bracket (8 left, 8 right, top→bottom).
- `adv` is the team code the player has advancing. For a decisive predicted
  score it's the higher-scoring team; for a predicted draw it's taken from the
  annotation (`(ALE)`, `1-1 CAN`, `pen Brasil`, …).

### `results.json` (written by the fetch Action)
```jsonc
{
  "lastUpdated": "...Z",
  "matches": {
    "r32-01": { "status": "FINISHED", "g1": 2, "g2": 1, "adv": "GER",
                "score": "2-1", "utcDate": "...Z", "source": "api" }
  },
  "lastFinishedId": "r32-03",
  "nextMatchId": "r32-04"
}
```

### `overrides.json`
Manual safety net: map a match id to `{ "g1": n, "g2": n, "adv": "XXX" }`.
Overrides always win over the API.

## The 16 matches (bracket order, from the Excel "Ronda Eliminatoria" block)

Left side: Alemania–Paraguay, Francia–Suecia, Sudáfrica–Canadá,
Países Bajos–Marruecos, Portugal–Croacia, España–Austria, EE.UU–Bosnia,
Bélgica–Senegal.
Right side: Brasil–Japón, C. Marfil–Noruega, México–Ecuador,
Inglaterra–RD Congo, Argentina–Cabo Verde, Australia–Egipto, Suiza–Argelia,
Colombia–Ghana.

## Build script — `scripts/build_predictions_r32.py`

Parses the `Ronda Eliminatoria` rows (78–93) of `Porra Mundial 2026.xlsx`:
- Splits the match label (`"Alemania - Paraguay"`) into two teams; maps the
  Spanish display names to team codes via a Spanish-name alias table.
- Parses each player cell `"g1-g2"` plus an optional advancer annotation. The
  annotation token (`ALE`, `MAR`, `pen Brasil`, `1-1 CAN`, …) is matched
  loosely against the two teams in that row (scoped, so fuzzy match is safe).
- Decisive scores set `adv` to the higher team automatically; draws require the
  annotation (warn if a drawn prediction is missing an advancer).
- Writes `RondaEliminatoria/data/predictions.json`.

## Results fetch

Extend the football-data.org fetch to the Round-of-32 stage: filter the WC
competition matches to the knockout stage, match each porra fixture by its team
pair, and record `g1`/`g2` (full-time) plus `adv` (API winner, falling back to
the penalties field). Writes `RondaEliminatoria/data/results.json`; runs from
the same scheduled GitHub Action.

Caveat: the 16 pairings are the porra's *guesses*. If the real Round of 32 pairs
teams differently, that fixture stays `pending` until a matching real result
exists. `overrides.json` covers any gap.

## Page UI (`index.html`, self-contained CSS+JS, dark theme reused)

1. **Header** — title "Porra Mundial 2026 · Ronda Eliminatoria", live badge,
   progress chip `X / 16 jugados`, last-updated chip.
2. **Leaderboard** — ranked list of the 15 players by total points; click a name
   to select. Selected row highlighted. Default selection = leader. Podium
   accents reused from the group-stage style.
3. **Bracket** — mirrors the reference image:
   - Center card: selected player's **name + total points** (+ rank).
   - 8 R32 boxes per side; decorative empty inner boxes for QF/SF; connector
     lines. Each R32 box shows both teams (flag + name) and the player's
     predicted score, the box tinted by points earned.
   - **0–3 color scale** with a legend: `0` muted red · `1` amber · `2` green ·
     `3` bright green with glow. The numeric point value also shows as a small
     badge per box so the count is unambiguous.
   - **Responsive:** on phones the side-by-side bracket reflows to a single
     stacked column of the 16 match boxes (same content + colors); the center
     card becomes a header above it.
4. **"Resultados Formato Excel" popup** — collapsible at the bottom (same toggle
   pattern as the group-stage site). Table: 16 matches (rows) × 15 players
   (columns); each cell = predicted score, colored 0–3; plus a column for the
   real result.

All points are computed client-side from `predictions.json` + `results.json`,
refreshed on an interval like the group-stage page.

## Out of scope (v1)

The Excel's bonus questions (Ganadores, Pichichi, Mejor Gol, Goles en Propia,
Más goleada, # Banda). Can be added later.

## Non-goals / constraints

- No build tooling or JS dependencies on the page — single static HTML file,
  matching the existing site.
- Do not modify the root group-stage `index.html` or its data.
