"""Syndicate 3123 (Lloyd's) RDS — the syndicate's own return, reported separately
from the IG (group) RDS.

The IG return and the S3123 return share the SAME damage factors, RPF bands and
per-risk loss formula. They differ in three places, all handled here:

  1. SHARE      — S3123 takes its consortium sub-share of each CONSORTIUM risk:
                  s3123_share = per_sc * s3123_factor   (the 12.5/30, 15/30,
                  10/30 split already computed by the engine), gated by
                  consortium eligibility — non-consortium layers have NO S3123
                  share at all.
                  FIX(recon 2026Q1, D5): this gate was missing; every layer got
                  a share, so 2022 pre-consortium lines (SXM-10, NUSANTARA LIMA)
                  leaked into the return — +1.56m on PF, +1.22m on GD, and
                  SXM-10 wrongly became the syndicate's Max Risk ($16.7m; the
                  correct pick is INMARSAT 6-F1, $12.5m).
  2. SELECTION  — Lloyd's "realistic disaster" caps rather than full-portfolio
                  aggregates:
                    Proton Flare   : all on-risk GEO            (uncapped)
                    Space Weather  : worst Lloyd's BUS TYPE, top-N spacecraft
                    Generic Defect : top-N individual GEO/MEO risks by loss
                    Space Debris   : LEO x RPF (time-decay, not the IG flat DR)
  3. NETTING    — S3123 cedes a 20% QS back to IG (the SAME cession that lands
                  in the IG number as s3123_qs; excluded spacecraft retain
                  100%), plus the consortium Agg XoL (30m max any one event).
                  Net = gross x 0.80 for non-excluded lines.
                  NOTE the two different exclusion concepts:
                    - NOT consortium (no share at all)      -> share gate (1)
                    - consortium but retains 100% (no QS)   -> qs_excluded (3)
                  INMARSAT 6-F1 / SPAINSAT NG-1 / AMAZONAS NEXUS are the latter.

Population & the Lloyd's bus type come from the Lloyd's SQL views
(vw_SpaceRDS_All_Lloyds_RDS / vw_SpaceRDS_SpaceWeather_Lloyds); see ingest.
OPEN (recon 2026Q1, O1): when run on the IG population, the schedule-derived
share ties JJ's Lloyd's book in aggregate (PF to £3) but differs per line from
the actual signed S3123 lines, which moves the GD top-10 selection (~1m) and
leaves Space Debris on a different method (share x RPF here vs JJ's
altitude-grouped basis). Ingesting per-line shares from the Lloyd's view, and
the SD method, are to be agreed with JJ — reconcile() itemises the residuals.
This module implements the CLEAN Lloyd's structure — it does not replicate the
manual workbook's per-quarter manual adjustments or known errors.
"""
from __future__ import annotations
from datetime import date, datetime
import pandas as pd

SCEN_ORDER = ["Proton Flare", "Space Weather", "Generic Defect",
              "Space Debris", "Max Risk"]

# FIX(recon 2026Q1, D5): mirror of engine.CONSORTIUM_START — kept local so this
# module stays importable standalone; config s3123_rds.consortium_start wins.
CONSORTIUM_START = date(2024, 7, 1)


def _as_date(x):
    if x is None:
        return None
    if isinstance(x, str):
        return datetime.fromisoformat(x).date()
    if isinstance(x, (pd.Timestamp, datetime)):
        return x.date()
    return x


# --------------------------------------------------------------------------- #
# config access (params.raw["s3123_rds"]) with safe defaults
# --------------------------------------------------------------------------- #
def _cfg(p) -> dict:
    c = (p.raw.get("s3123_rds") or {}) if hasattr(p, "raw") else {}
    sc = c.get("scenarios", {})
    # FIX(recon 2026Q1, O2): honor `active: false` — the 2026Q1 config carried
    # agg_xol {limit: 30m, attach: 0, active: false} and the old code ignored
    # `active` and applied attach=0, recovering the ENTIRE net: that is where
    # the shipped book's $0 nets came from.
    xol = c.get("agg_xol") or {}
    if isinstance(xol, dict) and not xol.get("active", True):
        xol = {}
    return {
        # 20% QS ceded back to IG (accepts the old key name for compatibility)
        "qs_to_ig": c.get("qs_to_ig", c.get("travelers_qs", 0.20)),
        "qs_excluded": set(c.get("qs_to_ig_excluded",
                                 c.get("travelers_qs_excluded", [])) or []),
        "agg_xol": xol,
        "consortium_start": _as_date(c.get("consortium_start")) or CONSORTIUM_START,
        "share_source": c.get("share", "consortium_factor"),
        "pf":  sc.get("proton_flare",  {"orbits": ["GEO-GSO"], "loss": 0.05}),
        "sw":  sc.get("space_weather", {"top_n": 4, "loss": 1.00}),
        "gd":  sc.get("generic_defect", {"top_n": 10,
                       "orbits": ["GEO-GSO", "MEO"], "loss": 0.50,
                       "apply_rpf": True}),
        "sd":  sc.get("space_debris",  {"orbits": ["LEO"], "apply_rpf": True}),
        "max_risk": c.get("max_risk", True),
    }


def s3123_share(df: pd.DataFrame, cfg: dict | None = None) -> pd.Series:
    """Per-layer S3123 share of the signed line, gated by consortium eligibility.

    Two share sources (config s3123_rds.share):
      * "consortium_factor" (default): per_sc * s3123_factor — the period
        schedule (12.5/30, 15/30, 10/30).
      * "lloyds_view": the syndicate's ACTUAL signed line per layer, ingested
        from vw_SpaceRDS_All_Lloyds_RDS (df["lloyds_signed_share"], attached
        by ingest). Falls back to the schedule for layers the view lacks.
        This is the O1 resolution path — aggregates coincide but per-line
        values move the GD top-10; switch after agreeing the source with JJ.

    FIX(recon 2026Q1, D5): previously ungated — every layer received a share.
    """
    share = df["per_sc"].astype(float) * df["s3123_factor"].astype(float)
    if (cfg or {}).get("share_source") == "lloyds_view" \
            and "lloyds_signed_share" in df.columns:
        lv = pd.to_numeric(df["lloyds_signed_share"], errors="coerce")
        # the view's line is per layer; split across a layer's spacecraft
        occ = pd.to_numeric(df.get("layer_occ", 1), errors="coerce").fillna(1)
        lv = lv / occ
        share = lv.where(lv.notna(), share)
    cons_start = (cfg or {}).get("consortium_start", CONSORTIUM_START)
    if "is_consortium" in df.columns:
        eligible = df["is_consortium"].fillna(False).astype(bool)
    else:  # raw frame — fall back to the inception floor alone
        eligible = pd.Series(True, index=df.index)
    incep = df["inception"].map(_as_date)
    eligible = eligible & (incep >= cons_start)
    return share.where(eligible, 0.0)


# --------------------------------------------------------------------------- #
# netting: per-line QS back to IG (excluded spacecraft retain 100%), then a
# scenario-aggregate Agg XoL (max recovery any one event).
# --------------------------------------------------------------------------- #
def _net(gross_by_line: pd.Series, df: pd.DataFrame, cfg: dict) -> float:
    qs = cfg["qs_to_ig"]
    excl = cfg["qs_excluded"]
    ceded = gross_by_line * df.loc[gross_by_line.index, "spacecraft_id"].map(
        lambda s: 0.0 if s in excl else qs)
    net = (gross_by_line - ceded).sum()
    xol = cfg["agg_xol"]
    # FIX(recon 2026Q1, O2): only apply the consortium Agg XoL when BOTH terms
    # are positive. With attach missing/0 the old code recovered the whole net
    # (recovery = min(net, limit)) and every scenario printed $0 net — a silent
    # config error, not a real result. A limit-only structure must be entered
    # with its true attachment.
    if xol:
        attach = float(xol.get("attach") or 0.0)
        limit = float(xol.get("limit") or 0.0)
        if attach > 0 and limit > 0:
            recovery = max(min(net - attach, limit), 0.0) if net > attach else 0.0
            net -= recovery
        else:
            import warnings
            warnings.warn("s3123: agg_xol ignored — set BOTH attach and limit "
                          "(> 0) in config; attach=0 would zero every net.")
    return float(net)


# --------------------------------------------------------------------------- #
# the four scenarios — each returns (gross, net, detail, picked-line index)
# --------------------------------------------------------------------------- #
def _proton(df, share, cfg):
    s = cfg["pf"]
    scope = df["orbit"].isin(s["orbits"]) & (df["on_risk_flag"] == 1)
    line = (share * s["loss"])[scope]
    return line.sum(), _net(line, df, cfg), "All GEO", line.index


def _space_weather(df, share, cfg):
    s = cfg["sw"]
    flag = df["on_risk_flag"] == 1
    bt = df["lloyds_bus_type"]
    valid = flag & bt.notna() & ~bt.isin(["", "None"])
    agg = (share[valid]).groupby(bt[valid]).sum()
    if not len(agg):
        return 0.0, 0.0, None, df.index[:0]
    worst = agg.idxmax()
    cand = share[valid & bt.eq(worst)].sort_values(ascending=False)
    top = cand.iloc[: s.get("top_n", 4)]
    line = top * s["loss"]
    return line.sum(), _net(line, df, cfg), worst, line.index


def _generic_defect(df, share, cfg):
    s = cfg["gd"]
    scope = df["orbit"].isin(s["orbits"]) & (df["on_risk_flag"] == 1)
    loss = share * s["loss"]
    if s.get("apply_rpf", True):
        loss = loss * df["rpf"]
    cand = loss[scope].sort_values(ascending=False)
    top = cand.iloc[: s.get("top_n", 10)]
    return top.sum(), _net(top, df, cfg), f"Top {s.get('top_n', 10)} risks", top.index


def _space_debris(df, share, cfg):
    # Lloyd's LEO debris = share x RPF (time-decay), NOT the IG flat 0.40 DR.
    # OPEN (recon 2026Q1, O1): JJ's 2026Q1 workings use an altitude-grouped
    # flat basis instead — method to be agreed before this is treated as final.
    s = cfg["sd"]
    scope = df["orbit"].isin(s["orbits"]) & (df["on_risk_flag"] == 1)
    # FIX(recon 2026Q1): untangled the one-line conditional (same behaviour).
    if s.get("apply_rpf", True):
        line = (share * df["rpf"])[scope]
    else:
        line = (share * s.get("loss", 1.0))[scope]
    return line.sum(), _net(line, df, cfg), "LEO group", line.index


def _max_risk(df, share, cfg):
    flag = df["on_risk_flag"] == 1
    g = share[flag].groupby(df.loc[flag, "spacecraft_name"]).sum()
    g = g[g > 0]                      # FIX(recon 2026Q1): ignore zero-share names
    if not len(g):
        return 0.0, 0.0, None, df.index[:0]
    biggest = g.idxmax()
    line = share[flag & df["spacecraft_name"].eq(biggest)]
    return line.sum(), _net(line, df, cfg), biggest, line.index


_CALC = {
    "Proton Flare": _proton, "Space Weather": _space_weather,
    "Generic Defect": _generic_defect, "Space Debris": _space_debris,
    "Max Risk": _max_risk,
}


def s3123_grid(df: pd.DataFrame, p) -> pd.DataFrame:
    """S3123 RDS by scenario — gross + net, at the syndicate's share."""
    if "lloyds_bus_type" not in df.columns:
        df = df.copy()
        df["lloyds_bus_type"] = None
    cfg = _cfg(p)
    share = s3123_share(df, cfg)     # FIX(recon 2026Q1, D5): eligibility-gated
    rows = []
    for scen in SCEN_ORDER:
        if scen == "Max Risk" and not cfg["max_risk"]:
            continue
        gross, net, detail, _ = _CALC[scen](df, share, cfg)
        rows.append({"entity": "S3123", "scenario": scen, "detail": detail,
                     "gross": float(gross), "net": float(net)})
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------- #
# reconciliation against the manual Lloyd's submission (optional QA)
# --------------------------------------------------------------------------- #
def reconcile(grid: pd.DataFrame, manual: dict) -> pd.DataFrame:
    """manual = {scenario: {'gross': x, 'net': y}}. Returns a diff frame."""
    out = []
    for _, r in grid.iterrows():
        m = manual.get(r["scenario"], {})
        out.append({
            "scenario": r["scenario"],
            "auto_gross": r["gross"], "manual_gross": m.get("gross"),
            "d_gross": (r["gross"] - m["gross"]) if m.get("gross") is not None else None,
            "auto_net": r["net"], "manual_net": m.get("net"),
            "d_net": (r["net"] - m["net"]) if m.get("net") is not None else None,
            "why": m.get("note", ""),
        })
    return pd.DataFrame(out)
