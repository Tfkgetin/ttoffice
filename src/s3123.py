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
def _cfg(p, key: str = "s3123_rds") -> dict:
    c = (p.raw.get(key) or {}) if hasattr(p, "raw") else {}
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


def s3123_share(df: pd.DataFrame, cfg: dict | None = None,
                factor_col: str = "s3123_factor") -> pd.Series:
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
    if factor_col not in df.columns:      # e.g. config without an S2126 schedule
        return pd.Series(0.0, index=df.index)
    share = df["per_sc"].astype(float) * df[factor_col].astype(float)
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
# the four scenarios — each returns (gross, net, detail, per-line gross Series)
# The 4th element is the per-layer GROSS contribution over the scenario's
# picked set (empty elsewhere); s3123_layer_contribs turns it into Per-Layer
# columns so the Lloyd's Summary can write Gross/Net as live =SUM() formulas.
# --------------------------------------------------------------------------- #
def _proton(df, share, cfg):
    s = cfg["pf"]
    scope = df["orbit"].isin(s["orbits"]) & (df["on_risk_flag"] == 1)
    line = (share * s["loss"])[scope]
    return line.sum(), _net(line, df, cfg), "All GEO", line


def _space_weather(df, share, cfg):
    # JJ's filed basis (vw_SpaceRDS_SpaceWeather_Lloyds + the All-RDS view):
    # for each Lloyd's bus-type group take its TOP-4 gross exposures, and the
    # Design-Deficiency RDS is the MAX of those top-4 sums across bus types —
    # i.e. the bus type whose four largest satellites sum highest (not the one
    # with the largest whole-fleet aggregate).
    s = cfg["sw"]
    flag = df["on_risk_flag"] == 1
    bt = df["lloyds_bus_type"]
    valid = flag & bt.notna() & ~bt.isin(["", "None"])
    if not valid.any():
        return 0.0, 0.0, None, pd.Series(dtype=float)
    n = s.get("top_n", 4)
    loss = s.get("loss", 1.0)
    best = None
    for grp, idx in df[valid].groupby(bt[valid]).groups.items():
        top = share.loc[idx].sort_values(ascending=False).head(n)
        tot = float((top * loss).sum())
        if best is None or tot > best[0]:
            best = (tot, top.index, grp)
    line = share.loc[best[1]] * loss
    return line.sum(), _net(line, df, cfg), best[2], line


def _generic_defect(df, share, cfg):
    s = cfg["gd"]
    scope = df["orbit"].isin(s["orbits"]) & (df["on_risk_flag"] == 1)
    loss = share * s["loss"]
    if s.get("apply_rpf", True):
        loss = loss * df["rpf"]
    if s.get("type") == "largest_manufacturer":
        # JJ's filed basis: 50% of the LARGEST MANUFACTURER's whole fleet,
        # policy-period (RPF) adjusted — not the top-N individual risks. Largest
        # = the manufacturer with the biggest RPF-weighted S3123 loss.
        by = loss[scope].groupby(df.loc[scope, "bus_manufacturer"]).sum()
        by = by[by > 0]
        if not len(by):
            return 0.0, 0.0, None, pd.Series(dtype=float)
        worst = by.idxmax()
        idx = df.index[scope & df["bus_manufacturer"].eq(worst)]
        return (loss[idx].sum(), _net(loss[idx], df, cfg),
                f"Largest manufacturer: {worst}", loss[idx])
    # default: top-N individual risks
    cand = loss[scope].sort_values(ascending=False)
    top = cand.iloc[: s.get("top_n", 10)]
    return top.sum(), _net(top, df, cfg), f"Top {s.get('top_n', 10)} risks", top


def _space_debris(df, share, cfg):
    # JJ's filed basis (vw_SpaceRDS_Space_Debris_Lloyds): 100% of all LEO
    # satellites in the SAME orbit range (RPF-adjusted), taking the MAX across
    # altitude groups — Group 1 = 400-800 km, Group 2 = 1200-1600 km. The
    # per-line loss is share x RPF (100% loss ratio); the RDS is the worst
    # single group. Needs altitude_km (ingest attaches it from the Lloyd's
    # attributes query); falls back to all-LEO when altitude is absent.
    s = cfg["sd"]
    flag = df["on_risk_flag"] == 1
    leo = df["orbit"].isin(s["orbits"]) & flag
    line = (share * df["rpf"]) if s.get("apply_rpf", True) \
        else (share * s.get("loss", 1.0))
    groups = s.get("altitude_groups")
    if groups and "altitude_km" in df.columns:
        alt = pd.to_numeric(df["altitude_km"], errors="coerce")
        best = None
        for grp in groups:
            m = leo & alt.ge(grp["low"]) & alt.lt(grp["high"])
            tot = float(line[m].sum())
            if best is None or tot > best[0]:
                best = (tot, df.index[m], grp["name"])
        if best is not None and len(best[1]):
            idx = best[1]
            return line[idx].sum(), _net(line[idx], df, cfg), f"LEO {best[2]}", line[idx]
    # fallback: all LEO (no altitude / no groups configured)
    return line[leo].sum(), _net(line[leo], df, cfg), "LEO (all — no altitude)", line[leo]


def _max_risk(df, share, cfg):
    flag = df["on_risk_flag"] == 1
    g = share[flag].groupby(df.loc[flag, "spacecraft_name"]).sum()
    g = g[g > 0]                      # FIX(recon 2026Q1): ignore zero-share names
    if not len(g):
        return 0.0, 0.0, None, pd.Series(dtype=float)
    biggest = g.idxmax()
    line = share[flag & df["spacecraft_name"].eq(biggest)]
    return line.sum(), _net(line, df, cfg), biggest, line


_CALC = {
    "Proton Flare": _proton, "Space Weather": _space_weather,
    "Generic Defect": _generic_defect, "Space Debris": _space_debris,
    "Max Risk": _max_risk,
}

# scenario -> short slug for Per Layer contribution column names. Only the four
# JJ RDS (not Max Risk) feed the Lloyd's Summary, so only these get columns.
SCEN_SLUG = {"Proton Flare": "pf", "Space Weather": "sw",
             "Generic Defect": "gd", "Space Debris": "sd"}


def s3123_layer_contribs(df: pd.DataFrame, p, cfg_key: str = "s3123_rds",
                         factor_col: str = "s3123_factor",
                         prefix: str = "s3123") -> pd.DataFrame:
    """Per-layer GROSS and NET contribution to each syndicate RDS, indexed like
    `df`. The scenario SELECTION is baked in: a layer outside a scenario's
    picked set is 0. Summing a column reproduces s3123_grid's gross/net for
    that scenario, so the Lloyd's Summary can render Gross/Net as live
    =SUM(Per Layer) formulas (the same backbone the IG Summary uses).

    Columns: {prefix}_{slug}_g (gross) and {prefix}_{slug}_n (net) for each of
    the four JJ scenarios (pf/sw/gd/sd).

    Net = gross per line, less the QS cession to IG (excluded spacecraft retain
    100%). This decomposition is exact ONLY while there is no active Agg XoL
    (a scenario-aggregate structure can't be split per line); when one is
    active the net columns are omitted and the Summary falls back to the
    engine's netted value.
    """
    if "lloyds_bus_type" not in df.columns:
        df = df.copy()
        df["lloyds_bus_type"] = None
    cfg = _cfg(p, cfg_key)
    share = s3123_share(df, cfg, factor_col)
    qs = cfg["qs_to_ig"]
    excl = cfg["qs_excluded"]
    net_decomposable = not cfg["agg_xol"]
    cede = df["spacecraft_id"].map(lambda s: 0.0 if s in excl else qs).astype(float)
    out = pd.DataFrame(index=df.index)
    for scen, slug in SCEN_SLUG.items():
        _, _, _, line = _CALC[scen](df, share, cfg)  # per-line gross Series
        g = pd.Series(0.0, index=df.index)
        if len(line):
            g.loc[line.index] = line.astype(float)
        out[f"{prefix}_{slug}_g"] = g
        if net_decomposable:
            out[f"{prefix}_{slug}_n"] = g * (1.0 - cede)
    return out


def s3123_grid(df: pd.DataFrame, p, cfg_key: str = "s3123_rds",
               factor_col: str = "s3123_factor",
               label: str = "S3123") -> pd.DataFrame:
    """Syndicate RDS by scenario — gross + net, at the syndicate's consortium
    share. Parameterised so the SAME methodology serves both S3123 and S2126:
      * cfg_key    — config block (s3123_rds / s2126_rds)
      * factor_col — per-layer sub-share column (s3123_factor / s2126_factor)
      * label      — grid entity tag (S3123 / S2126)
    S2126 differs only by config: 5/30 factor and qs_to_ig=0 (no cession to IG,
    so net = gross). Same scenarios, RPF, eligibility gate."""
    if "lloyds_bus_type" not in df.columns:
        df = df.copy()
        df["lloyds_bus_type"] = None
    cfg = _cfg(p, cfg_key)
    share = s3123_share(df, cfg, factor_col)  # FIX(recon 2026Q1, D5): eligibility-gated
    rows = []
    for scen in SCEN_ORDER:
        if scen == "Max Risk" and not cfg["max_risk"]:
            continue
        gross, net, detail, _ = _CALC[scen](df, share, cfg)
        rows.append({"entity": label, "scenario": scen, "detail": detail,
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
