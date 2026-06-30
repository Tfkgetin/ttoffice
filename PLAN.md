# Space RDS Pipeline — Automation Plan

**Objective:** Replace the manual quarterly workbook process with a reproducible, parameterised, validated Python pipeline that ingests the SQL extract and produces all RDS outputs — while reconciling exactly against the existing workbook so nothing is lost in translation.

---

## 1 · What the workbook does today (as reverse-engineered from vTT)

```
SQL (vw_SpaceRDS_OnRisk, ~174 rows)
   │  paste into Input Data B19:AA193
   ▼
ENGINE (Input Data, computed columns P–DD)
   ├─ Time decay        P: RPF banded on months to off-risk (0.2/0.4/0.6/0.8/1.0)
   │                    AC/AD: LEO-specific months-left + debris RPF
   ├─ Damage ratios     AE: Space Debris DR by orbit (GEO 5% / MEO 10% / LEO 40%)
   ├─ Equity            AG/AH: IG Equity % by inception year × S3123 consortium factor
   ├─ Per-S/C           AK = exposure ÷ LayerID occurrence count
   ├─ External RI       AO–AX: 10 outwards slots allocated by inception window → AY
   ├─ S3123 inwards     BD/BE: consortium, non-excluded → 20% × consortium factor
   ├─ IGR QS            BG/BH/BI: Entity+Mapping_Code+UWYear lookup (FUL 50%, FIID 85%)
   └─ IGR XoL per layer BJ/BK/BL: layer-level breach of entity deductible
   ▼
SCENARIOS (per layer)
   ├─ Proton Flare      GEO-GSO × 5%                       (BY–CD)
   ├─ Space Weather     mfr-group totals of GEO per-S/C    (CI–CN, max picked in pivot)
   ├─ Generic Defect    (GEO|MEO) × 50% × RPF              (CR–CW)
   ├─ Space Debris      (GEO|MEO) × 100% × orbit DR        (CY–DD)
   └─ Max Risk          per-layer totals of BF / BL        (BT–BW, max picked in pivot)
   ▼
AGGREGATION
   ├─ Summary           entity × scenario netting grid + FIG long format
   ├─ Bus Manufacturer  worst-mfr pivot per entity (feeds Space Weather row)
   ├─ Spacecraft        largest per-layer pivot (feeds Max Risk row)
   ├─ Netting Waterfall scenario-level waterfall per entity (Gross→ExtQS→IGR QS→IGR XoL→Net)
   └─ Changes           Q-on-Q vs From_PQ
```

## 2 · Pipeline goals

| Goal | How |
|---|---|
| Remove the manual paste | Ingest directly from SQL (pyodbc) or from a CSV export |
| Kill formula fragility | All logic in version-controlled Python, parameters in YAML |
| Prove correctness | **Reconciliation mode**: run the engine on the workbook's own raw columns and diff every computed column + every Summary figure against the workbook values |
| Quarterly turnaround | One command: `python run_pipeline.py --config config/2026Q1.yaml` |
| Q-on-Q automatic | Pipeline persists each quarter's results; Changes computed from stored prior |
| Outputs people recognise | Netting waterfalls, summary grid, pivots — exported as CSVs + a formatted Excel |

## 3 · Build phases (dynamic workflow: validate each stage before the next)

| Phase | Deliverable | Validation gate |
|---|---|---|
| 1. Parameters | `config/2026Q1.yaml` capturing every treaty/scenario parameter | Visual diff vs param tabs |
| 2. Ingest | Raw layer dataframe (cols B–AA) from workbook or SQL CSV | Row count + column dtypes |
| 3. Engine | All per-layer computed columns | Per-column diff vs workbook (tolerance $1) |
| 4. Scenarios | 5 scenario loss columns per layer | Diff vs BY–DD blocks |
| 5. Netting | Scenario-level waterfalls per entity | Diff vs Netting Waterfall tab + Summary |
| 6. Aggregation | Summary grid, pivots, FIG long format | Diff vs Summary / Bus Manufacturer / Spacecraft |
| 7. Outputs | CSVs + Excel + JSON run manifest | recalc clean, no errors |

## 4 · Architecture

```
space_rds_pipeline/
├── config/
│   └── 2026Q1.yaml          # ALL parameters — one file per quarter
├── src/
│   ├── ingest.py            # raw layer data from workbook / CSV / SQL
│   ├── parameters.py        # typed parameter loading + validation
│   ├── engine.py            # per-layer computed columns (the Engine)
│   ├── scenarios.py         # 5 scenario loss calculators
│   ├── netting.py           # scenario-level entity waterfalls
│   ├── aggregate.py         # summary grid, pivots, FIG long format
│   ├── validate.py          # reconciliation vs the workbook
│   └── outputs.py           # CSV / Excel / manifest export
├── run_pipeline.py          # CLI orchestrator
└── output/                  # per-run timestamped results
```

**Design principles**
- Parameters never live in code. The YAML is the single source of truth per quarter; next quarter = copy the YAML, update dates/terms.
- The engine is pure dataframe-in → dataframe-out. No I/O inside calculation modules.
- Reconciliation is a first-class mode, not an afterthought: `--reconcile` runs everything against the workbook and prints a per-column / per-figure diff table.
- SQL ingest is pluggable: same pipeline accepts `--source workbook|csv|sql`.

## 5 · Parameter inventory (extracted from vTT workbook)

| Parameter | Value | Source tab |
|---|---|---|
| As-at date | 2026-04-01 | Input Data AC15 |
| RPF bands (months→factor) | 0→0.2, 6→0.4, 12→0.6, 18→0.8, 999→1.0 | Input Data O10:P14 |
| Proton Flare insured loss | 5% (GEO-GSO only) | Summary P9 |
| Generic Defect loss | 50% (GEO+MEO, × RPF) | Summary P3 |
| Space Debris loss | 100% (GEO+MEO, × orbit DR) | Summary P6 |
| Debris DR by orbit | LEO 40%, MEO 10%, GEO-GSO 5% | Summary U4:Y6 |
| LEO debris RPF bands | 0→0.2, 6→0.4, 12→0.6, 18→0.8 (shifted +1) | Summary R4:S7 |
| IG Equity Share by inception year | 2024: 9.9%, 2025: 7.5%, 2026: 7.5% | IG Equity Share |
| S3123 consortium factor by date | 2022-11: 41.67%, 2026-01: 50%, 2026-04: 33.3% | S3123 Consortium Share |
| S3123 QS cession | 20%, window 2022-11-01→2026-12-31 | S3123 QS |
| S3123 exclusions | spacecraft 14294, 13257, 10590 | S3123 QS |
| IGR QS | FUL 50%, FIID 85% (all codes/years), FIBL 0 | IGR QS |
| IGR XoL | FUL $245m xs $50m · FIID $72.5m xs $7.5m | IGR XoL |
| Outwards RI slots (10) | windows + cessions 11.8%/3.8%/20%×6 | Input Data AO6:AX11 |

## 6 · Risks / open questions

- **AI (LayerID OCC)** semantics assumed = row count per Program+Layer key; confirmed in reconciliation.
- **Space Weather XoL** is applied at scenario-aggregate level in Summary; per-layer BJ is only for Max Risk. Pipeline implements both.
- **Other Ext RI (BC)** currently always 0; implemented as a hook so future structures drop in.
- **#REF in RDS_date named range** — workbook wart; pipeline takes the date from config.
- **LEO exclusion in Space Debris** is replicated as-is (formula excludes LEO despite DR table including it) and flagged in validation output for methodology review.

---

## 7 · Validation results (build of 2026-06-04, vs vTT workbook)

**Per-layer engine — 18 columns, zero diff on all 169 rows:**
rpf, leo_months_left, leo_debris_rpf, debris_dr, equity_pct, equity_usd, layer_occ,
exposure_usd, per_sc, ext_qs, net_ext_qs, s3123_qs, net_ext, igr_qs_rate,
igr_qs_ceded, net_of_qs, xol_ceded, net_of_xol — **ALL OK**

**Per-layer scenarios — 10 columns, zero diff** (Space Debris compared on the
GEO+MEO subset matching the Input Data CY block; Summary-level total includes LEO).

**Headline figures — 20 checks vs Summary + Netting Waterfall, all exact ($1 tol):**
every entity × scenario gross, the FUL & FIID Space Weather waterfalls
(incl. FIID XoL 9,928,305 → net 7,500,000 and FUL's no-XoL 49,043,092),
FIID Proton Flare net 2,149,690, Generic Defect net 4,825,479, Max Risk picks
(SXM-10 / SPAINSAT NG-1).

**Discoveries made during reconciliation (now encoded in config):**
1. IGR QS is genuinely (entity, code, year)-specific: FUL 2021–22 = 40% for
   ASC/ASL/ASR, ASO 2022 only (ASO 2021 stays 50%).
2. `Is Consortium` (AF) is a data attribute — FALSE for spacecraft 17163/17950 —
   driving both equity share and S3123 eligibility.
3. Space Weather = worst manufacturer's **entire fleet** (all orbits), ranked by
   net-of-QS per entity; per-layer CI block (GEO-only) is a legacy view.
4. Space Debris Summary includes LEO at 40% DR with **no time decay**; the
   AD (LEO debris RPF) column exists but is unused in the Summary formula —
   worth a methodology query.
5. Max Risk keys on **spacecraft name** (sums across layers covering the same
   spacecraft), not on layer.

## 8 · Usage

```bash
# Reconciliation mode (proves parity vs the workbook)
python run_space_rds.py --config config/2026Q1.yaml --reconcile

# Production (point ingest.source at csv/sql in the YAML)
python run_space_rds.py --config config/2026Q2.yaml

# Or install once (pip install -e .) and run from anywhere:
space-rds --config config/2026Q1.yaml --reconcile
```

Outputs land in `output/<quarter>/`: per_layer.csv, netting_waterfalls.csv,
summary_grid.csv, space_weather_by_manufacturer.csv, max_risk_by_layer.csv,
reconciliation.csv, manifest.json.

**Next quarter:** copy the YAML, update as-at date + any treaty changes,
switch `ingest.source` to `sql`, run. The reconciliation harness stays useful:
run the workbook in parallel for a quarter or two, diff, then retire the paste.
