"""The Engine: per-layer computed columns, mirroring Input Data P-BL exactly."""
from __future__ import annotations
import math
from datetime import date, datetime
import pandas as pd

# FIX(recon 2026Q1, D5): the consortium did not exist before this date. Used as a
# hard floor on S3123 QS / equity eligibility so a wrongly-restated is_consortium
# flag in the live view (e.g. spacecraft 17163 SXM-10, 17950 NUSANTARA LIMA —
# 2022 inceptions flagged TRUE in 2026Q2 data) cannot leak pre-consortium
# business into the IG add-ons. Overridable via config s3123_qs.consortium_start.
CONSORTIUM_START = date(2024, 7, 1)


def _as_date(x):
    """Normalise str / datetime / Timestamp / date to datetime.date."""
    if x is None:
        return None
    if isinstance(x, str):
        return datetime.fromisoformat(x).date()
    if isinstance(x, (pd.Timestamp, datetime)):
        return x.date()
    return x


def _months_between(d1, d2) -> int:
    """Excel DATEDIF(d1, d2, "M") — complete months."""
    if d2 is None or d1 is None:
        return 0
    m = (d2.year - d1.year) * 12 + (d2.month - d1.month)
    if d2.day < d1.day:
        m -= 1
    return m


def run_engine(df: pd.DataFrame, p) -> pd.DataFrame:
    df = df.copy()
    as_at = p.as_at

    # Pipeline-side exclusion of consortium placing-basis layers.
    # The SQL view only filters 'Consortium Declaration'; plain 'Consortium'
    # slips through. Rather than alter shared SQL, drop them here when the
    # config asks. Excluded rows are tagged so the Excluded sheet can show them.
    df["excluded_reason"] = None
    excl_bases = [b.strip().lower() for b in
                  (p.raw.get("exclude_placing_basis") or [])]
    if excl_bases and "placing_basis" in df.columns:
        pb = df["placing_basis"].astype(str).str.strip().str.lower()
        hit = pb.isin(excl_bases)
        df.loc[hit, "excluded_reason"] = "Consortium placing basis (pipeline filter)"
        _exc = df[hit].copy()
        # give the excluded set a per_sc figure for reporting (exposure / occ)
        _exc["per_sc"] = pd.to_numeric(_exc["layer_signed_exposure"],
                                       errors="coerce").fillna(0.0)
        run_engine.last_excluded = _exc
        df = df[~hit].copy()
    else:
        run_engine.last_excluded = df.iloc[0:0].copy()

    # Re-apply the on-risk date filter. SQL does this at source, but a data
    # correction (e.g. fixed off-risk date) can make a layer off-risk after
    # ingest — drop those so corrected expiries take effect.
    if "off_risk_date" in df.columns:
        def _onrisk(d):
            o = d["off_risk_date"]
            on = d.get("on_risk_date")
            if o is None:
                return True  # keep undated (handled elsewhere)
            if o < as_at:
                return False
            if on is not None and on > as_at:
                return False
            return True
        keep = df.apply(_onrisk, axis=1)
        if (~keep).any():
            _off = df[~keep].copy()
            _off["per_sc"] = pd.to_numeric(_off["layer_signed_exposure"],
                                           errors="coerce").fillna(0.0)
            prev = run_engine.last_excluded
            _off["excluded_reason"] = "Off-risk at as-at (after data correction)"
            run_engine.last_excluded = pd.concat([prev, _off], ignore_index=True) \
                if len(prev) else _off
            df = df[keep].copy()

    # Exclude NULL-orbit rows from ALL calculations (not just orbit-gated ones).
    # These have no Seradata ID/orbit; they must not leak into Max Risk, Portfolio
    # totals, or any sum. Captured (with original exposure) for the Excluded tab.
    # On by default; set exclude_null_orbit: false in config to keep them.
    if p.raw.get("exclude_null_orbit", True) and "orbit" in df.columns:
        orbit_s = df["orbit"].astype(str)
        null_orbit = orbit_s.isin(["NULL", "None", "nan", ""]) | df["orbit"].isna()
        if null_orbit.any():
            _no = df[null_orbit].copy()
            if "per_sc" not in _no.columns or _no["per_sc"].isna().all():
                _no["per_sc"] = pd.to_numeric(_no["layer_signed_exposure"],
                                              errors="coerce").fillna(0.0)
            _no["excluded_reason"] = "No orbit (no Seradata ID — placeholder/unlaunched)"
            prev = run_engine.last_excluded
            run_engine.last_excluded = pd.concat([prev, _no], ignore_index=True) \
                if len(prev) else _no
            df = df[~null_orbit].copy()

    # Re-added duplicate guard. When a programme is re-added via the manual add
    # table with sequential layer ids (1,2,3…), the base SQL view may still carry
    # the same spacecraft under its native (large) LayerId — double-counting that
    # spacecraft's exposure. Signature: a (program, spacecraft) pair holding BOTH
    # a small sequential id (<1000, a manual add) AND a large native id (>=1000,
    # base view). Drop the base-view row, keep the manual add. Legitimate
    # multi-layer satellites use all-native ids, so they never match.
    if p.raw.get("dedupe_readded", True) and {"program_id", "spacecraft_id",
                                              "layer_id"} <= set(df.columns):
        lid = pd.to_numeric(df["layer_id"], errors="coerce")
        SEQ_MAX = 1000
        stale_idx = []
        for (pid, scid), grp in df.groupby(["program_id", "spacecraft_id"]):
            ids = pd.to_numeric(grp["layer_id"], errors="coerce")
            has_seq = (ids < SEQ_MAX).any()
            has_native = (ids >= SEQ_MAX).any()
            if has_seq and has_native:
                # drop the native base-view row(s); keep the sequential add
                stale_idx.extend(grp.index[ids >= SEQ_MAX].tolist())
        if stale_idx:
            _dup = df.loc[stale_idx].copy()
            if "per_sc" not in _dup.columns or _dup["per_sc"].isna().all():
                _dup["per_sc"] = pd.to_numeric(_dup["layer_signed_exposure"],
                                               errors="coerce").fillna(0.0)
            _dup["excluded_reason"] = ("Re-added duplicate (base-view row "
                                       "superseded by manual Add Layer)")
            prev = run_engine.last_excluded
            run_engine.last_excluded = pd.concat([prev, _dup], ignore_index=True) \
                if len(prev) else _dup
            df = df.drop(index=stale_idx).copy()

    # On-risk flag: workbook column Q is 1 for live rows (extract is pre-filtered)
    df["on_risk_flag"] = 1

    # --- Time decay -------------------------------------------------------
    # P: RPF banded on DATEDIF(as_at, off_risk_date, "M")
    df["rpf"] = df["off_risk_date"].map(
        lambda d: p.rpf(_months_between(as_at, d)) if d else 0.0)

    # AC: LEO months left = ROUNDUP((off_risk - as_at)/30.5) * flag, floored 0
    def _leo_months(row):
        if row["orbit"] != "LEO" or row["off_risk_date"] is None:
            return 0
        m = math.ceil((row["off_risk_date"] - as_at).days / 30.5) * row["on_risk_flag"]
        return max(m, 0)
    df["leo_months_left"] = df.apply(_leo_months, axis=1)

    # AD: LEO debris RPF (shifted band lookup)
    df["leo_debris_rpf"] = df["leo_months_left"].map(p.leo_debris_rpf)

    # AE: Space Debris damage ratio by orbit
    dmg = p.scenarios["space_debris"]["damage_by_orbit"]
    df["debris_dr"] = df["orbit"].map(lambda o: dmg.get(o, 0.0))

    # --- Per-spacecraft economics ------------------------------------------
    # FIX(recon 2026Q1, fix #7): moved ABOVE the consortium/equity block so that
    # equity can be computed on a per-S/C basis (it previously used the whole
    # layer's exposure_usd, which would overcount equity by the spacecraft count
    # on a multi-S/C consortium layer; no 2026Q1 impact — all consortium layers
    # are single-spacecraft — but wrong going forward).
    # AI: LayerID OCC = number of spacecraft rows sharing the layer
    df["layer_occ"] = df.groupby("layer_key")["layer_key"].transform("size")
    df["exposure_usd"] = df["layer_signed_exposure"].astype(float)
    # AK: Per S/C
    df["per_sc"] = df["exposure_usd"] / df["layer_occ"] * df["on_risk_flag"]

    # --- Consortium / equity ----------------------------------------------
    # AF: Is Consortium — ingested attribute (False for a handful of
    # non-consortium layers, e.g. spacecraft 17163 / 17950 in 2026Q1)
    df["is_consortium"] = df["is_consortium"].fillna(False).astype(bool)

    # FIX(recon 2026Q1, D5): hard floor — nothing incepting before the consortium
    # start can be consortium business, whatever the ingested flag says. The
    # 2026Q2 live view restated is_consortium = TRUE for scids 17163 / 17950
    # (2022 inceptions), which put spurious S3123 QS into FIHL (+312.5k PF/SD,
    # +916.7k GD, +3.33m Max Risk) and SXM-10 into the S3123 return.
    q = p.s3123_qs
    cons_start = _as_date(q.get("consortium_start")) or CONSORTIUM_START
    incep_d = df["inception"].map(_as_date)
    pre_consortium = incep_d < cons_start
    df["is_consortium"] = df["is_consortium"] & ~pre_consortium

    # AG / AH
    df["equity_pct"] = df.apply(
        lambda r: p.equity_pct(r["inception"], r["is_consortium"]), axis=1)
    df["s3123_factor"] = df["inception"].map(p.s3123_factor)
    # S2126 consortium sub-share (new participant from 2026-04-01, 5/30); 0 before
    # then / when the config carries no S2126 schedule. Used ONLY by the separate
    # S2126 syndicate RDS — S2126 has NO QS to IG, so it does not touch the IG
    # book (no s2126_qs / equity add-back).
    df["s2126_factor"] = df["inception"].map(p.s2126_factor)
    # FIX(recon 2026Q1, fix #7): per_sc basis (was exposure_usd) — mirrors s3123_qs.
    df["equity_usd"] = df["per_sc"] * df["equity_pct"] * df["s3123_factor"]

    # --- External (outwards) RI ---------------------------------------------
    def _ext_qs(row):
        total = 0.0
        for slot in p.outwards_slots:
            if slot["from"] <= row["inception"] <= slot["to"]:
                total += row["per_sc"] * slot["pct"]
        return total
    df["ext_qs"] = df.apply(_ext_qs, axis=1)
    df["other_ext_ri"] = 0.0                       # BC — hook for future structures
    df["net_ext_qs"] = df["per_sc"] - df["ext_qs"]  # BB
    df["net_ext"] = df["net_ext_qs"] - df["other_ext_ri"]  # BF

    # --- S3123 inwards QS ----------------------------------------------------
    # FIX(recon 2026Q1, D5): the window's lower bound is floored at
    # consortium_start (config previously opened it at 2022-11-01; the earliest
    # legitimate S3123 layer incepts 2024-12-15, so the floor is loss-free).
    # NOTE: the excluded list means "retains 100%, no QS ceded back to IG" —
    # those spacecraft are still consortium (they still carry equity); do NOT
    # use it to fix wrong is_consortium flags.
    date_from = max(_as_date(q["date_from"]), cons_start)
    date_to = _as_date(q["date_to"])
    df["s3123_eligible"] = df.apply(
        lambda r: bool(r["is_consortium"])
        and r["spacecraft_id"] not in q["excluded"]
        and date_from <= _as_date(r["inception"]) <= date_to,
        axis=1)

    def _s3123(row):
        if not row["s3123_eligible"]:
            return 0.0
        return row["per_sc"] * q["cession"] * row["s3123_factor"]
    df["s3123_qs"] = df.apply(_s3123, axis=1)      # BE

    # --- IGR QS ---------------------------------------------------------------
    df["uw_year"] = df["inception"].map(lambda d: d.year if d else None)
    df["igr_qs_rate"] = df.apply(
        lambda r: p.igr_qs_rate(r["entity"], r["mapping_code"], r["uw_year"]), axis=1)
    df["igr_qs_ceded"] = df["net_ext"] * df["igr_qs_rate"]   # BH
    df["net_of_qs"] = df["net_ext"] - df["igr_qs_ceded"]     # BI

    # --- IGR XoL (per layer) ---------------------------------------------------
    df["xol_ceded"] = df.apply(
        lambda r: p.xol_recovery(r["entity"], r["net_of_qs"]), axis=1)  # BJ
    df["total_fibl_ceded"] = df["igr_qs_ceded"] + df["xol_ceded"]       # BK
    df["net_of_xol"] = df["net_of_qs"] - df["xol_ceded"]                # BL

    return df
