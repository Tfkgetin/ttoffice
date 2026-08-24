#!/usr/bin/env python3
"""Gross RDS by class code (Mapping_Code), straight off per_layer.csv.

    python rds_by_class.py                    # latest run under output/
    python rds_by_class.py --as-at 2026-07-01
    python rds_by_class.py --entity FUL       # operating-entity perspective

Writes RDS_by_class_<as_at>.csv and .xlsx next to the run, and prints the table.
"""
from __future__ import annotations
import argparse
from pathlib import Path
import pandas as pd

SCEN = [("pf", "Proton Flare"), ("gd", "Generic Defect"), ("sd", "Space Debris"),
        ("sw", "Space Weather"), ("mr", "Max Risk")]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--output-root", default="output")
    ap.add_argument("--as-at", help="run folder, e.g. 2026-07-01 (default: latest)")
    ap.add_argument("--entity", default="fihl",
                    help="perspective: fihl (group, default), ful, fiid")
    a = ap.parse_args()

    root = Path(a.output_root)
    runs = sorted(d for d in root.iterdir() if (d / "per_layer.csv").exists())
    if not runs:
        raise SystemExit(f"No <as_at>/per_layer.csv under {root}/")
    run = next((d for d in runs if d.name == a.as_at), None) if a.as_at else runs[-1]
    if run is None:
        raise SystemExit(f"{a.as_at} not found. Available: "
                         + ", ".join(d.name for d in runs))

    df = pd.read_csv(run / "per_layer.csv", low_memory=False)
    if "mapping_code" not in df.columns:
        raise SystemExit("per_layer.csv has no mapping_code column")
    df["mapping_code"] = df["mapping_code"].fillna("(blank)").astype(str)

    ent = a.entity.lower()
    cols, missing = {}, []
    for slug, label in SCEN:
        c = f"{slug}_{ent}"
        if c in df.columns:
            cols[label] = pd.to_numeric(df[c], errors="coerce").fillna(0)
        else:
            missing.append(c)
    if not cols:
        raise SystemExit(f"No {ent} scenario columns found in per_layer.csv")
    if missing:
        print(f"  ! not in this run, omitted: {', '.join(missing)}")

    work = pd.DataFrame({"mapping_code": df["mapping_code"], **cols})
    work["Gross exposure"] = pd.to_numeric(df["per_sc"], errors="coerce").fillna(0)

    out = work.groupby("mapping_code").sum(numeric_only=True)
    out.insert(0, "Layers", df.groupby("mapping_code").size())
    out.loc["TOTAL"] = out.sum()

    order = ["Layers", "Gross exposure"] + [l for _, l in SCEN if l in out.columns]
    out = out[[c for c in order if c in out.columns]]

    print(f"\nGross RDS by class code — {run.name}, {ent.upper()} perspective\n")
    with pd.option_context("display.float_format", lambda v: f"{v:,.0f}",
                           "display.width", 200):
        print(out.to_string())

    csv_p = run / f"RDS_by_class_{run.name}.csv"
    out.to_csv(csv_p)
    try:
        xl_p = run / f"RDS_by_class_{run.name}.xlsx"
        with pd.ExcelWriter(xl_p, engine="openpyxl") as xw:
            out.to_excel(xw, sheet_name="RDS by class")
        print(f"\nWritten: {csv_p}\n         {xl_p}")
    except Exception as e:
        print(f"\nWritten: {csv_p}   (xlsx skipped: {e})")


if __name__ == "__main__":
    main()
