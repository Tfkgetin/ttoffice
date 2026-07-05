"""Reconciliation: diff every pipeline figure against the workbook's own values."""
from __future__ import annotations
import pandas as pd

TOL = 1.0  # dollars

ENGINE_CHECKS = [
    ("rpf", "wb_rpf"),
    ("leo_months_left", "wb_months_left"),
    ("leo_debris_rpf", "wb_leo_debris_rpf"),
    ("debris_dr", "wb_debris_dr"),
    ("equity_pct", "wb_equity_pct"),
    ("equity_usd", "wb_equity_usd"),
    ("layer_occ", "wb_layer_occ"),
    ("exposure_usd", "wb_exposure_usd"),
    ("per_sc", "wb_per_sc"),
    ("ext_qs", "wb_ext_qs"),
    ("net_ext_qs", "wb_net_ext_qs"),
    ("s3123_qs", "wb_s3123_qs"),
    ("net_ext", "wb_net_ext"),
    ("igr_qs_rate", "wb_igr_qs_rate"),
    ("igr_qs_ceded", "wb_igr_qs_ceded"),
    ("net_of_qs", "wb_net_of_qs"),
    ("xol_ceded", "wb_xol_ceded"),
    ("net_of_xol", "wb_net_of_xol"),
]

SCENARIO_CHECKS = [
    ("pf_fihl", "wb_pf_fihl"), ("pf_ful", "wb_pf_ful"),
    ("pf_fiid", "wb_pf_fiid"), ("pf_fibl_ceded", "wb_pf_fibl_ceded"),
    ("gd_fihl", "wb_gd_fihl"), ("gd_ful", "wb_gd_ful"),
    ("gd_fiid", "wb_gd_fiid"),
    ("sd_fihl", "wb_sd_fihl"), ("sd_ful", "wb_sd_ful"),
    ("sd_fiid", "wb_sd_fiid"),
]


# Engine vs the SQL view's own computed columns (present in SQL-ingest mode)
SQL_CHECKS = [
    ("igr_qs_rate", "sql_igr_qs_rate"),
    ("ext_qs", "sql_ext_qs"),
    ("net_ext_qs", "sql_net_ext_qs"),
    ("igr_qs_ceded", "sql_igr_qs_ceded"),
    ("net_of_qs", "sql_net_of_qs"),
    ("xol_ceded", "sql_xol_ceded"),
    ("total_fibl_ceded", "sql_total_fibl_ceded"),
    ("net_of_xol", "sql_net_of_xol"),
]


def reconcile_columns(df: pd.DataFrame, checks=None) -> pd.DataFrame:
    """Per-column reconciliation. Returns summary with max abs diff + n mismatches."""
    checks = checks or (ENGINE_CHECKS + SCENARIO_CHECKS)
    out = []
    for ours, theirs in checks:
        if theirs not in df.columns or ours not in df.columns:
            out.append({"column": ours, "status": "MISSING", "max_diff": None, "n_bad": None})
            continue
        a = pd.to_numeric(df[ours], errors="coerce").fillna(0.0)
        b = pd.to_numeric(df[theirs], errors="coerce").fillna(0.0)
        diff = (a - b).abs()
        n_bad = int((diff > TOL).sum())
        out.append({
            "column": ours,
            "status": "OK" if n_bad == 0 else "MISMATCH",
            "max_diff": float(diff.max()),
            "n_bad": n_bad,
            "ours_total": float(a.sum()),
            "wb_total": float(b.sum()),
        })
    return pd.DataFrame(out)


def reconcile_value(name: str, ours: float, theirs: float, tol=TOL) -> dict:
    d = abs(ours - theirs)
    return {"figure": name, "ours": ours, "workbook": theirs,
            "diff": d, "status": "OK" if d <= tol else "MISMATCH"}
