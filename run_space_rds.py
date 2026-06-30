#!/usr/bin/env python3
"""Space RDS pipeline — orchestrator.

Usage:
    python run_pipeline.py --config config/2026Q1.yaml [--outdir output/2026Q1]
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from src.parameters import Params
from src import ingest, engine, scenarios, netting as net_mod, validate, outputs


# Headline targets from the workbook (Summary + Netting Waterfall tabs).
# Used in reconciliation mode to prove the pipeline reproduces the workbook.
WB_HEADLINES = [
    # (entity, scenario, field, workbook value)
    ("FIHL", "Proton Flare",   "gross", 47_706_969),
    ("FIHL", "Space Weather",  "gross", 267_843_602),
    ("FIHL", "Generic Defect", "gross", 133_349_351),
    ("FIHL", "Space Debris",   "gross", 108_771_648),
    ("FIHL", "Max Risk",       "gross", 39_999_960),
    ("FUL",  "Proton Flare",   "gross", 29_792_884),
    ("FUL",  "Proton Flare",   "net",   12_217_154),
    ("FUL",  "Space Weather",  "gross", 122_607_729),
    ("FUL",  "Space Weather",  "net",   49_043_092),
    ("FUL",  "Generic Defect", "gross", 93_137_025),
    ("FUL",  "Space Debris",   "gross", 75_341_155),
    ("FIID", "Proton Flare",   "gross", 17_914_085),
    ("FIID", "Proton Flare",   "net",   2_149_690),
    ("FIID", "Space Weather",  "gross", 145_235_873),
    ("FIID", "Space Weather",  "xol_ceded", 9_928_305),
    ("FIID", "Space Weather",  "net",   7_500_000),
    ("FIID", "Generic Defect", "gross", 40_212_326),
    ("FIID", "Generic Defect", "net",   4_825_479),
    ("FIID", "Space Debris",   "gross", 29_430_523),
    ("FIID", "Max Risk",       "gross", 28_857_627),
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", required=True)
    ap.add_argument("--outdir", default=None)
    ap.add_argument("--as-at", dest="as_at", default=None,
                    help="override the config's as_at_date (YYYY-MM-DD)")
    ap.add_argument("--reconcile", action="store_true",
                    help="validate computed columns + headlines against the workbook")
    args = ap.parse_args()

    p = Params.load(args.config)
    if args.as_at:
        import datetime as _dt
        p.as_at = _dt.date.fromisoformat(args.as_at)
    outdir = args.outdir or f"output/{p.as_at}"

    print(f"Space RDS pipeline · {p.quarter} · as-at {p.as_at}")

    print("[1/6] Ingest…")
    df = ingest.load(p)
    print(f"      {len(df)} rows · {df['layer_key'].nunique()} layers")

    print("[2/6] Engine…")
    df = engine.run_engine(df, p)

    sql_cols = [c for c in df.columns if c.startswith("sql_")]
    if sql_cols:
        import pandas as pd
        chk = [(a, b) for a, b in validate.SQL_CHECKS if b in df.columns]
        sq = validate.reconcile_columns(df, chk)
        nb = (sq["status"] != "OK").sum()
        print(f"      engine vs SQL view: "
              f"{'ALL OK' if nb == 0 else f'{nb} MISMATCHED'} ({len(sq)} checks)")
        if nb:
            print(sq[sq["status"] != "OK"].to_string(index=False))

    print("[3/6] Scenarios…")
    per_layer, sw, mr = scenarios.run_scenarios(df, p)

    print("[4/6] Netting + summary grid…")
    nets = net_mod.entity_scenario_netting(per_layer, p)
    grid = net_mod.summary_grid(per_layer, p)

    print("[5/6] Reconciliation…")
    import pandas as pd
    if args.reconcile:
        sub = per_layer[per_layer["orbit"] != "LEO"]
        sd = [(a, b) for a, b in validate.SCENARIO_CHECKS if a.startswith("sd_")]
        rest = validate.ENGINE_CHECKS + [(a, b) for a, b in validate.SCENARIO_CHECKS
                                         if not a.startswith("sd_")]
        recon = pd.concat([validate.reconcile_columns(per_layer, rest),
                           validate.reconcile_columns(sub, sd)], ignore_index=True)
        n_bad = (recon["status"] != "OK").sum()
        print(f"      per-layer columns: "
              f"{'ALL OK' if n_bad == 0 else f'{n_bad} MISMATCHED'} "
              f"({len(recon)} checks)")

        hl = []
        for ent, scen, field, target in WB_HEADLINES:
            row = grid[(grid["entity"] == ent) & (grid["scenario"] == scen)]
            ours = float(row.iloc[0][field]) if len(row) else float("nan")
            hl.append(validate.reconcile_value(f"{ent} · {scen} · {field}", ours, target))
        hl = pd.DataFrame(hl)
        bad = hl[hl["status"] != "OK"]
        print(f"      headline figures:  "
              f"{'ALL OK' if len(bad) == 0 else f'{len(bad)} MISMATCHED'} "
              f"({len(hl)} checks)")
        if len(bad):
            print(bad.to_string(index=False))
        recon = pd.concat([recon, hl.rename(columns={"figure": "column"})],
                          ignore_index=True)
    else:
        recon = pd.DataFrame([{"column": "reconcile", "status": "SKIPPED"}])

    print("[6/6] Export…")
    out = outputs.export(outdir, per_layer, sw, mr, nets, grid, recon, p)
    print(f"Done → {out}")


if __name__ == "__main__":
    main()
