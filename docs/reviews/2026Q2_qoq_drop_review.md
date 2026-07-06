# 2026Q2 Quarter-on-Quarter Drop — Review & Checks

**Question raised:** the 2026Q2 run shows a large drop vs 2026Q1 — is anything wrong?

**Verdict:** No calculation error. The drop is **genuine run-off**, correctly
computed. Two items on the run *look* alarming but are red herrings, and one item
is a real thing to verify (July renewals). Details below.

---

## Scope of this review

| | |
|---|---|
| Prior (Q1) book | `Space_RDS_results_2026-04-01.xlsx` (as-at 2026-04-01) |
| Current (Q2) book | `Space_RDS_results_2026-07-01.xlsx` (as-at 2026-07-01) |
| Run command | `python run_space_rds.py --config config\2026Q2.yaml --reconcile` |
| Prior comparison basis | `prior_workbook: output/2026-04-01/Space_RDS_results_2026-04-01.xlsx` (frozen Q1 auto book) |

---

## Checks performed

| # | Check | Method | Result |
|---|-------|--------|--------|
| 1 | Layer-count reconciliation | Per Layer row counts + new/dropped from Changes tab | **Consistent** — 153 − 59 + 18 = 112 ✓ |
| 2 | Re-keying artifact? | Overlap of new vs dropped spacecraft names | **None** — true turnover, not a key glitch |
| 3 | Are the drops legitimate expiries? | `off_risk_date` of dropped spacecraft in the Q1 book | **Yes** — all expire ≤ 2026-06, before the 1-Jul as-at |
| 4 | Does the expiry profile explain "59 dropped"? | `off_risk_date` distribution by month (Q1 book) | **Yes** — 7 (Apr) + 13 (May) + 39 (Jun) = 59 exactly |
| 5 | Entity-level loss movement | Summary gross/net, Q1 vs Q2, per entity × scenario | Real; magnitudes follow the run-off (see below) |
| 6 | Reconciliation MISMATCHes | Traced `--reconcile` targets in `run_space_rds.py` | **Red herring** — targets are hard-coded Q1 values |
| 7 | Prior-baseline warning | `prior_seed` uncached S3123/equity warning | Affects the QoQ *matrix* only, not the current book |

---

## Finding 1 — The drop is real run-off, and it reconciles exactly

The book shrinks **153 → 112 layers**. The layers expiring **before** the 1-Jul
as-at (from `off_risk_date` in the Q1 book):

| Expiry month | Layers |
|---|---|
| 2026-04 | 7 |
| 2026-05 | 13 |
| 2026-06 | **39** |
| **≤ Jun total** | **59** |

`153 − 59 dropped + 18 new = 112` ✓. The engine is correctly dropping policies
whose cover ended before 1 Jul. Big movers spot-checked:

| Spacecraft | Entity | Inception | Off-risk | Note |
|---|---|---|---|---|
| Arabsat BADR 8 | FUL | 2025-07-01 | 2026-06-30 | 1-yr policy, ran off |
| EUTELSAT 36D | FIID | 2025-07-01 | 2026-06-30 | 1-yr policy, ran off |
| E10B / E13F / E13G | FIID | 2025-07-01 | 2026-06-30 | 1-yr policies, ran off |
| SPAINSAT NG-1 | FIID | 2024-12/2025-01 | 2026-05-30 | ran off |
| AMAZONAS NEXUS | FIID | 2025-06-07 | 2026-06-06 | ran off |
| SXM-10 | FUL | 2022-12-31 | 2026-06-07 | ran off |

No overlap between the *new* and *dropped* spacecraft lists → genuine turnover
(SXM-10 out, SXM-11/12 in; INTELSAT 41/43/44/45 new), **not** a re-keying artifact.

### Entity-level gross movement (Q1 → Q2)

| Entity | Q1 gross | Q2 gross | Δ% |
|---|--:|--:|--:|
| FIHL | 661.2m | 411.3m | −37.8% |
| FUL | 359.0m | 320.8m | −10.6% |
| FIID | 268.8m | 132.9m | −50.5% |
| FIBL | 374.9m | 194.9m | −48.0% |

FIID falls hardest — consistent with it being the most-ceded, running-off entity.
Max Risk (single largest spacecraft) is flat, as expected.

---

## Finding 2 — The real thing to verify: **July renewals**

Dropped exposure **606m** vs new business only **181m**. That June-30 cliff is a
classic **1-July annual renewal tranche** (many policies incept 2025-07-01 →
expire 2026-06-30). If the Q2 run was taken *at* as-at 1 Jul, those policies have
expired but their renewals (incept 2026-07-01) may not be **bound / entered in the
source yet** — so the book shows a hole that fills once renewals come in.

- This is a **data-timing** question, not a code bug — the engine's boundary logic
  *does* include a renewal incepting exactly on the as-at date.
- **Action:** confirm with underwriting that the 1-July renewals of the expiring
  tranche are in the source. If not, the Q2 book understates the real position.

**Secondary:** Q1 flagged *11 layers with placeholder expiry dates ("source-date
overrides pending true expiries")*. Confirm none of the 59 drops are
placeholder-driven (which would be a premature, wrong drop).

---

## Finding 3 — Two red herrings (safe to ignore for Q2)

**A. The 19 reconciliation MISMATCHes are meaningless for Q2.**
`--reconcile` compares against `WB_HEADLINES` in `run_space_rds.py`, which are
**hard-coded Q1 filed values** (e.g. FIHL Space Weather = `300_953_877`). Running
it on Q2 compares Q2 vs Q1 targets, so everything mismatches by construction.
*Fix:* drop `--reconcile` for Q2, or make the targets quarter-aware.

**B. The prior-baseline warning distorts only the QoQ matrix, not the current book.**
`prior_seed: auto prior book carries no cached S3123/equity` means the Q1 book was
saved with formulas but **no cached values**, so the prior's FIHL add-ons read as 0
→ the Changes-tab movement for FIHL is overstated. *Fix:* open the Q1 book in Excel
and re-save once (forces a recalc so values cache), or point `prior_workbook` at a
persisted run. The current Q2 numbers are unaffected.

---

## Recommended follow-ups (optional automation)

1. **Renewal-gap check** — flag expiring layers with no matching renewal, so the
   July-trough effect is explicit rather than inferred.
2. **Quarter-aware `--reconcile`** — stop comparing Q2 against Q1 targets.
3. **Placeholder-date report** — surface the 11 override-dated layers each quarter.

---

*Prepared from the two run books and the 2026Q2 console log. All figures are USD,
per-spacecraft signed exposure. Scenarios are single-event RDS views and are never
summed.*
