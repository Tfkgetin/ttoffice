#!/usr/bin/env python3
"""Space RDS pipeline — orchestrator.

Usage:
    python run_space_rds.py --config config/2026Q1.yaml [--outdir output/2026Q1]
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from src.parameters import Params
from src import ingest, engine, scenarios, netting as net_mod, validate, outputs


# Headline targets in --reconcile mode.
# FIX(recon 2026Q1): these are the CORRECTED-AUTO 2026Q1 figures — the shipped
# book adjusted for the one pipeline bug (D5: spurious S3123 eligibility on
# SXM-10 / NUSANTARA LIMA, since fixed in engine/s3123). They were proven
# equal to Manual + D1..D4 (EUTELSAT 36D timing, VIASAT off-risk correction,
# MEASAT inception restatement, live exposure drift) to the dollar — see
# Space_RDS_2026Q1_Auto_vs_Manual_Bridge.xlsx.
# NOTE: valid only against the 2026Q1 data vintage (as-at 2026-04-01, live view
# as of 2026-07-02, 11 data_corrections active). A later run against restated
# live data may legitimately drift — that is data, not logic. FIHL grosses are
# COMBINED (incl S3123 QS + IG equity — the manual Change Narrative / Exec
# basis, headline across the book); the ex-add-on split is in the Summary memo
# columns and on the "S3123 & Equity (IG)" tab. FIBL is not targeted (receiver
# basis, one hand-keyed manual cell on Max Risk).
# The built-in headline targets are the FILED 2026Q1 figures. They are only
# meaningful for a 2026Q1 run; other quarters must supply their own via the
# config key `reconcile_headlines` (list of {entity, scenario, field, target}),
# or the headline check is skipped rather than comparing against the wrong quarter.
WB_HEADLINES_QUARTER = "2026Q1"
WB_HEADLINES = [
    # (entity, scenario, field, corrected 2026Q1 value)
    ("FIHL", "Proton Flare",   "gross",  52_806_166),
    ("FIHL", "Space Weather",  "gross", 300_953_877),
    ("FIHL", "Generic Defect", "gross", 146_248_706),
    ("FIHL", "Space Debris",   "gross", 121_353_137),
    ("FIHL", "Max Risk",       "gross",  39_999_960),
    ("FUL",  "Proton Flare",   "gross",  29_792_854),
    ("FUL",  "Proton Flare",   "net",    12_217_141),
    ("FUL",  "Space Weather",  "gross", 122_607_603),
    ("FUL",  "Space Weather",  "net",    49_043_041),
    ("FUL",  "Generic Defect", "gross",  91_281_234),   # carries the VIASAT correction
    ("FUL",  "Space Debris",   "gross",  75_341_125),
    ("FUL",  "Max Risk",       "gross",  39_999_960),
    ("FIID", "Proton Flare",   "gross",  18_188_860),   # carries EUTELSAT 36D
    ("FIID", "Proton Flare",   "net",     2_182_663),
    ("FIID", "Space Weather",  "gross", 150_731_383),
    ("FIID", "Space Weather",  "xol_ceded", 10_587_766),
    ("FIID", "Space Weather",  "net",     7_500_000),
    ("FIID", "Generic Defect", "gross",  41_311_428),
    ("FIID", "Generic Defect", "net",     4_957_371),
    ("FIID", "Space Debris",   "gross",  29_705_298),
    ("FIID", "Max Risk",       "gross",  28_857_627),
]


def _drop_manual(df):
    """Rows minus manual inclusions — used to keep manual layers (which have no
    SQL-view columns) out of the engine-vs-view reconciliation."""
    if "manual_include" in df.columns:
        return df[~df["manual_include"].fillna(False).astype(bool)]
    return df


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", required=True)
    ap.add_argument("--outdir", default=None)
    ap.add_argument("--as-at", dest="as_at", default=None,
                    help="override the config's as_at_date (YYYY-MM-DD)")
    ap.add_argument("--reconcile", action="store_true",
                    help="validate computed columns + headlines against the workbook")
    ap.add_argument("--no-prior", action="store_true",
                    help="skip prior-quarter seeding even if the config sets prior_workbook")
    args = ap.parse_args()

    p = Params.load(args.config)
    if args.as_at:
        import datetime as _dt
        p.as_at = _dt.date.fromisoformat(args.as_at)
    outdir = args.outdir or f"output/{p.as_at}"

    print(f"Space RDS pipeline · {p.quarter} · as-at {p.as_at}")

    print("[1/7] Ingest…")
    df = ingest.load(p)
    print(f"      {len(df)} rows · {df['layer_key'].nunique()} layers")
    mi = getattr(ingest.load, "last_manual_includes", None) or []
    inc_ok = [c for c in mi if c.get("status") == "INCLUDED"]
    skipped = [c for c in mi if str(c.get("status", "")).startswith("SKIPPED")]
    if inc_ok:
        mx = sum(float(c.get("layer_signed_exposure") or c.get("per_sc") or 0)
                 for c in inc_ok)
        print(f"      + {len(inc_ok)} manual inclusion(s) (${mx:,.0f}) injected — "
              f"see Python Adjustments tab")
    if skipped:
        print(f"      · {len(skipped)} manual inclusion(s) skipped — renewal now "
              f"in source (double-count safeguard)")
    split_warn = ingest.check_consortium_split_coverage(p)
    if split_warn:
        print(f"      ⚠ {split_warn}")

    print("[2/7] Engine…")
    df = engine.run_engine(df, p)

    sql_cols = [c for c in df.columns if c.startswith("sql_")]
    if sql_cols:
        import pandas as pd
        chk = [(a, b) for a, b in validate.SQL_CHECKS if b in df.columns]
        # Manual inclusions are injected AFTER the SQL, so the view has no
        # computed columns for them — comparing engine vs view on those rows
        # always "mismatches" (NaN→0). Exclude them; they're audited on the
        # Python Adjustments tab, not reconciled against a view that never saw them.
        rec_df = _drop_manual(df)
        n_excl = len(df) - len(rec_df)
        sq = validate.reconcile_columns(rec_df, chk)
        nb = (sq["status"] != "OK").sum()
        excl_note = f" · {n_excl} manual-inclusion row(s) excluded" if n_excl else ""
        print(f"      engine vs SQL view: "
              f"{'ALL OK' if nb == 0 else f'{nb} MISMATCHED'} ({len(sq)} checks{excl_note})")
        if nb:
            print(sq[sq["status"] != "OK"].to_string(index=False))

    print("[3/7] Scenarios…")
    per_layer, sw, mr = scenarios.run_scenarios(df, p)

    print("[4/7] Netting + summary grid…")
    nets = net_mod.entity_scenario_netting(per_layer, p)
    grid = net_mod.summary_grid(per_layer, p)

    # ------------------------------------------------------------------ #
    # [5/7] Prior quarter — seed QoQ from the FROZEN manual workbook.
    # FIX(recon 2026Q1): the Changes / Exposure Bridge / Loss Movement
    # sheets shipped as "baseline" because no prior pipeline run exists.
    # Re-running the LIVE view with --as-at <prior close> is NOT a valid
    # prior (restatements pollute it — VIASAT, MEASAT, 36D would vanish
    # from movement). The frozen quarter-close workbook is the correct
    # filed prior; prior_seed reproduces it exactly (validated 2025Q4,
    # 10/10 rows). Once THIS run is persisted, switch to run-vs-run and
    # remove prior_workbook from config.
    # ------------------------------------------------------------------ #
    changes = None
    prior_path = (p.raw or {}).get("prior_workbook")
    if prior_path and not args.no_prior:
        print("[5/7] Prior quarter (frozen workbook)…")
        try:
            from src import prior_seed
            prior_layers = prior_seed.load_prior_workbook(prior_path)
            # Restore prior S3123/equity add-ons when the auto book cached them
            # as formulas — keeps FIHL combined on the same basis as current.
            prior_layers = prior_seed.restore_addons(prior_layers, per_layer, p)
            prior_grid = net_mod.summary_grid(prior_layers, p)
            label = (p.raw or {}).get("prior_as_at_label") or "prior close"
            changes = prior_seed.build_changes(per_layer, grid,
                                               prior_layers, prior_grid,
                                               prior_as_at=label)
            s = changes["layers"]["summary"]
            print(f"      prior: {len(prior_layers)} layers from workbook · "
                  f"{s['new_layers']} new / {s['dropped_layers']} dropped / "
                  f"{s['moved_layers']} moved this quarter")
            gap_n = s.get("renewal_gap_layers", 0)
            if gap_n:
                gap_x = s.get("renewal_gap_exposure", 0.0)
                print(f"      ⚠ renewal-gap: {gap_n} dropped layer(s) "
                      f"(${gap_x:,.0f} exposure) have NO current layer for the "
                      f"same spacecraft — verify these are true non-renewals, "
                      f"not renewals not yet bound/entered. See the Changes tab.")
        except Exception as e:
            print(f"      WARNING: prior seeding failed ({e}) — "
                  f"book will render as baseline")
            changes = None
    else:
        print("[5/7] Prior quarter… skipped (no prior_workbook in config)")

    print("[6/7] Reconciliation…")
    import pandas as pd
    if args.reconcile:
        # Exclude manual inclusions — they carry no SQL-view columns to reconcile
        # against (injected post-SQL); reconciling them would flag every check.
        pl_rec = _drop_manual(per_layer)
        n_excl = len(per_layer) - len(pl_rec)
        sub = pl_rec[pl_rec["orbit"] != "LEO"]
        sd = [(a, b) for a, b in validate.SCENARIO_CHECKS if a.startswith("sd_")]
        rest = validate.ENGINE_CHECKS + [(a, b) for a, b in validate.SCENARIO_CHECKS
                                         if not a.startswith("sd_")]
        recon = pd.concat([validate.reconcile_columns(pl_rec, rest),
                           validate.reconcile_columns(sub, sd)], ignore_index=True)
        n_bad = (recon["status"] != "OK").sum()
        excl_note = f" · {n_excl} manual-inclusion row(s) excluded" if n_excl else ""
        print(f"      per-layer columns: "
              f"{'ALL OK' if n_bad == 0 else f'{n_bad} MISMATCHED'} "
              f"({len(recon)} checks{excl_note})")

        # Headline figures reconcile ONLY against same-quarter filed targets.
        # Priority: config's `reconcile_headlines`; else the built-in set when
        # this run IS that quarter; else skip (comparing e.g. Q2 to Q1 targets
        # produces meaningless mismatches).
        cfg_hl = (p.raw or {}).get("reconcile_headlines")
        if cfg_hl:
            targets = [(h["entity"], h["scenario"], h["field"], float(h["target"]))
                       for h in cfg_hl]
            tgt_src = f"config · {p.quarter}"
        elif p.quarter == WB_HEADLINES_QUARTER:
            targets = WB_HEADLINES
            tgt_src = f"built-in · {WB_HEADLINES_QUARTER}"
        else:
            targets = None

        if targets:
            hl = []
            for ent, scen, field, target in targets:
                row = grid[(grid["entity"] == ent) & (grid["scenario"] == scen)]
                ours = float(row.iloc[0][field]) if len(row) else float("nan")
                hl.append(validate.reconcile_value(
                    f"{ent} · {scen} · {field}", ours, target))
            hl = pd.DataFrame(hl)
            bad = hl[hl["status"] != "OK"]
            print(f"      headline figures ({tgt_src}): "
                  f"{'ALL OK' if len(bad) == 0 else f'{len(bad)} MISMATCHED'} "
                  f"({len(hl)} checks)")
            if len(bad):
                print(bad.to_string(index=False))
            recon = pd.concat([recon, hl.rename(columns={"figure": "column"})],
                              ignore_index=True)
        else:
            print(f"      headline figures: skipped — no filed targets for "
                  f"{p.quarter} (built-in targets are {WB_HEADLINES_QUARTER}). "
                  f"Add `reconcile_headlines:` to the config to enable.")
    else:
        recon = pd.DataFrame([{"column": "reconcile", "status": "SKIPPED"}])

    print("[7/7] Export…")
    # `changes` is forwarded to outputs.export as a keyword so the report can
    # build the Changes tab + movement waterfalls. If outputs.export has not
    # yet been updated to accept it, fall back to the old signature and say so
    # (one-line fix in outputs.py: add `changes=None` to export() and pass it
    # through to excel_report.write_results(..., changes=changes, ...)).
    try:
        out = outputs.export(outdir, per_layer, sw, mr, nets, grid, recon, p,
                             changes=changes)
    except TypeError:
        if changes is not None:
            print("      NOTE: outputs.export() does not accept changes= yet — "
                  "prior-quarter sheets will render as baseline. Add "
                  "`changes=None` to outputs.export's signature and forward it "
                  "to excel_report.write_results.")
        out = outputs.export(outdir, per_layer, sw, mr, nets, grid, recon, p)
    print(f"Done → {out}")


if __name__ == "__main__":
    main()
