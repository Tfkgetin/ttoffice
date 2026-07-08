"""Seed the pipeline's quarter-on-quarter comparison from a FROZEN manual
workbook instead of a persisted prior pipeline run.

WHY (recon 2026Q1): the Changes / WF · Exposure Bridge / WF · Loss Movement
sheets activate only when a prior pipeline run is persisted — the 2026Q1 book
shipped as "baseline" because none existed. Re-running the pipeline with
--as-at <prior close> does NOT create a valid prior: the live view rewinds the
date filter only, so the "prior" would already contain every later restatement
(VIASAT-3 F2 off-risk, MEASAT inceptions, EUTELSAT 36D, exposure drift,
is_consortium flags) and genuine movement would silently net to zero.
The correct prior for a filed-vs-filed comparison is the frozen quarter-close
workbook — e.g. T:\\Actuarial\\Risk_and_Capital\\PMLs\\2025\\Q4\\Space\\
Space_RDS_2025Q4_template - including Equity Share.xlsx.

VALIDATED: the loader + the pipeline's own netting reproduce the filed 2025Q4
Summary exactly (10/10 FUL & FIID rows across the six-column cascade incl. both
non-zero XoL cells, 5/5 FIHL gross & net, equity block) and the filed 2026Q1
book to the dollar. See Space_RDS_2025Q4_parity_test.xlsx.

USAGE (in run.py, once per quarter until a persisted prior run exists):

    from prior_seed import load_prior_workbook, build_changes
    from netting import summary_grid

    prior_layers = load_prior_workbook(cfg["prior_workbook"])
    prior_grid   = summary_grid(prior_layers, p)      # same params: netting
                                                      # uses factors/XoL only
    changes = build_changes(per_layer, grid, prior_layers, prior_grid,
                            prior_as_at="2026-01-01 (frozen 2025Q4 manual)")
    excel_report.write_results(..., changes=changes, ...)

Once the 2026Q1 pipeline run is persisted, prior seeding from manual books can
be retired and QoQ becomes run-vs-run.
"""
from __future__ import annotations
import warnings
import numpy as np
import pandas as pd

DEBRIS_DR = {"GEO-GSO": 0.05, "MEO": 0.10, "LEO": 0.40}
EXT_QS_DEFAULT = 0.20   # validated: every in-force slot cedes 20% (Q4 & Q1)


# --------------------------------------------------------------------------- #
# loader
# --------------------------------------------------------------------------- #
def _find_header_row(path, sheet="Input Data", scan=40):
    probe = pd.read_excel(path, sheet_name=sheet, header=None, nrows=scan)
    for i in range(len(probe)):
        if (probe.iloc[i].astype(str).str.strip() == "ProgramId").any():
            return i
    raise ValueError(f"'ProgramId' header not found in first {scan} rows of "
                     f"'{sheet}' — is this a Space RDS template?")


def _col(df, *names, required=True):
    for n in names:
        if n in df.columns:
            return df[n]
    if required:
        raise KeyError(f"none of {names} present — template layout changed?")
    return None


def load_prior_workbook(path, sheet="Input Data") -> pd.DataFrame:
    """Parse a frozen prior-quarter workbook into the pipeline's per-layer
    schema. Handles three formats:
      - AUTO results book (has a 'Per Layer' tab) — preferred from 2026Q2 on:
        point prior_workbook at last quarter's pipeline output.
      - 2025Q4-style manual template: no per-layer External QS / S3123 columns;
        IGR rates on the 'IGR QS' tab by (Entity, Mapping_Code, Year).
      - 2026Q1-style manual template: per-layer 'External QS',
        'Inward S3123 QS', 'QS FIBL Ceded' columns (rates derived per layer).
    """
    sheets = pd.ExcelFile(path).sheet_names
    if "Per Layer" in sheets and sheet == "Input Data":
        return _load_prior_auto_book(path)
    hdr = _find_header_row(path, sheet)
    raw = pd.read_excel(path, sheet_name=sheet, header=hdr)
    raw.columns = [str(c).strip() for c in raw.columns]
    raw = raw.dropna(subset=["ProgramId"])
    per_sc = pd.to_numeric(_col(raw, "Per S/C"), errors="coerce")
    raw = raw[per_sc.fillna(0) > 0].copy()

    L = pd.DataFrame(index=raw.index)
    L["program_id"] = _col(raw, "ProgramId").astype(int)
    L["layer_id"] = _col(raw, "LayerId").astype(int)
    L["spacecraft_id"] = pd.to_numeric(_col(raw, "SpacecraftId"),
                                       errors="coerce").astype("Int64")
    L["spacecraft_name"] = _col(raw, "Spacecraft Name")
    L["entity"] = _col(raw, "Entity")
    L["mapping_code"] = _col(raw, "Mapping_Code", "Mapping Code",
                             required=False)
    L["orbit"] = _col(raw, "Orbit Category", "Orbit")
    L["bus_manufacturer"] = _col(raw, "Bus Manufacturer")
    L["inception"] = pd.to_datetime(_col(raw, "Inception"))
    L["off_risk_date"] = pd.to_datetime(_col(raw, "Off_Risk_Date",
                                             "Off Risk Date"))
    L["rpf"] = pd.to_numeric(_col(raw, "Risk_Period_Factor"), errors="coerce")
    L["per_sc"] = pd.to_numeric(_col(raw, "Per S/C"), errors="coerce")
    uw = _col(raw, "UW Year", required=False)
    L["uw_year"] = (pd.to_numeric(uw, errors="coerce")
                    if uw is not None else np.nan)
    L["uw_year"] = L["uw_year"].fillna(L["inception"].dt.year).astype(int)
    cons = _col(raw, "Is Consortium", required=False)
    L["is_consortium"] = (pd.to_numeric(cons, errors="coerce").fillna(0) > 0
                          if cons is not None else False)

    # --- external QS: per-layer column if the template carries it -----------
    ext = _col(raw, "External QS", required=False)
    L["ext_qs"] = (pd.to_numeric(ext, errors="coerce").fillna(0)
                   if ext is not None else L["per_sc"] * EXT_QS_DEFAULT)
    L["other_ext_ri"] = 0.0
    L["net_ext_qs"] = L["per_sc"] - L["ext_qs"]
    L["net_ext"] = L["net_ext_qs"] - L["other_ext_ri"]

    # --- IGR QS rate: per-layer derivation, else the 'IGR QS' rate tab ------
    ceded = _col(raw, "QS FIBL Ceded", required=False)
    if ceded is not None:
        c = pd.to_numeric(ceded, errors="coerce").fillna(0)
        L["igr_qs_rate"] = np.where(L["net_ext"] > 0, c / L["net_ext"], 0.0)
        L["igr_qs_ceded"] = c
    else:
        try:
            tab = pd.read_excel(path, sheet_name="IGR QS")
            tab.columns = [str(c).strip() for c in tab.columns]
            rate = {(r.Entity, r.Mapping_Code, int(r.Year)): float(r.Cession)
                    for r in tab.itertuples()}
        except Exception as e:
            raise ValueError("template has neither a 'QS FIBL Ceded' column "
                             "nor an 'IGR QS' rate tab") from e

        def _lookup(row):
            if row["entity"] not in ("FUL", "FIID"):
                return 0.0
            v = rate.get((row["entity"], row["mapping_code"], row["uw_year"]))
            if v is None:   # latest year on file for that entity + mapping
                cands = {k: val for k, val in rate.items()
                         if k[0] == row["entity"] and k[1] == row["mapping_code"]}
                v = cands[max(cands)] if cands else None
            if v is None:
                warnings.warn(f"prior_seed: no IGR rate for "
                              f"{row['entity']}/{row['mapping_code']}/"
                              f"{row['uw_year']} — using 0")
                v = 0.0
            return v
        L["igr_qs_rate"] = L.apply(_lookup, axis=1)
        L["igr_qs_ceded"] = L["net_ext"] * L["igr_qs_rate"]
    L["net_of_qs"] = L["net_ext"] - L["igr_qs_ceded"]

    # --- IG add-ons (zero in pre-S3123 templates) ----------------------------
    s3 = _col(raw, "Inward S3123 QS", required=False)
    L["s3123_qs"] = (pd.to_numeric(s3, errors="coerce").fillna(0)
                     if s3 is not None else 0.0)
    eq = _col(raw, "Equity Share USD", required=False)
    L["equity_usd"] = (pd.to_numeric(eq, errors="coerce").fillna(0)
                       if eq is not None else 0.0)

    # --- remaining engine-schema columns -------------------------------------
    L["debris_dr"] = L["orbit"].map(DEBRIS_DR).fillna(0.0)
    L["on_risk_flag"] = 1
    L["layer_key"] = (L["program_id"].astype(str) + "|"
                      + L["layer_id"].astype(str))
    occ = L.groupby("layer_key")["layer_key"].transform("size")
    L["layer_occ"] = occ
    L["exposure_usd"] = L["per_sc"] * occ
    L["xol_ceded"] = 0.0          # XoL is scenario-aggregate, not per-layer
    L["net_of_xol"] = L["net_of_qs"]
    L["excluded_reason"] = None

    if not len(L):
        raise ValueError("prior workbook parsed to zero on-risk layers")
    return L.reset_index(drop=True)


def _load_prior_auto_book(path) -> pd.DataFrame:
    """Load a prior AUTO results book via its Per Layer tab (pipeline-native
    columns). Numeric stage columns may ship as formulas (uncached) — anything
    unreadable is recomputed from raw fields with the validated logic."""
    L = pd.read_excel(path, sheet_name="Per Layer")
    L.columns = [str(c).strip() for c in L.columns]
    need = {"program_id", "layer_id", "spacecraft_id", "entity", "orbit",
            "per_sc", "inception", "off_risk_date"}
    missing = need - set(L.columns)
    if missing:
        raise ValueError(f"auto prior book missing columns: {missing}")
    L = L[pd.to_numeric(L["per_sc"], errors="coerce").fillna(0) > 0].copy()
    L["per_sc"] = L["per_sc"].astype(float)
    L["inception"] = pd.to_datetime(L["inception"])
    L["off_risk_date"] = pd.to_datetime(L["off_risk_date"])
    for c in ("bus_manufacturer", "spacecraft_name", "mapping_code"):
        if c not in L.columns:
            L[c] = None
    # recompute any stage column that arrived NaN (formula not cached)
    def _num(col, fallback):
        s = pd.to_numeric(L.get(col), errors="coerce") if col in L.columns \
            else pd.Series(np.nan, index=L.index)
        return s.fillna(fallback)
    L["ext_qs"] = _num("ext_qs", L["per_sc"] * EXT_QS_DEFAULT)
    L["other_ext_ri"] = _num("other_ext_ri", 0.0)
    L["net_ext_qs"] = L["per_sc"] - L["ext_qs"]
    L["net_ext"] = L["net_ext_qs"] - L["other_ext_ri"]
    rate = _num("igr_qs_rate", 0.0)
    L["igr_qs_rate"] = rate
    L["igr_qs_ceded"] = _num("igr_qs_ceded", L["net_ext"] * rate)
    L["net_of_qs"] = _num("net_of_qs", L["net_ext"] - L["igr_qs_ceded"])
    L["s3123_qs"] = _num("s3123_qs", 0.0)
    L["equity_usd"] = _num("equity_usd", 0.0)
    if float(L["s3123_qs"].sum()) == 0.0 and float(L["equity_usd"].sum()) == 0.0:
        warnings.warn(
            "prior_seed: auto prior book carries no cached S3123/equity "
            "values (formulas uncached) — the exec-basis matrix will show "
            "FIHL combined = FIHL ex for the prior. Recalc-save the prior "
            "book once, or persist runs with values, to restore the add-ons.")
    if "rpf" not in L.columns or pd.to_numeric(L["rpf"], errors="coerce").isna().any():
        # stage columns ship as uncached formulas — rebuild rpf from dates
        # using the validated DATEDIF band logic and the book's own as-at
        # (read from its Audit tab).
        import re
        from datetime import date
        as_at = None
        try:
            aud = pd.read_excel(path, sheet_name="Audit", header=None, nrows=5)
            for v in aud.values.ravel():
                m = re.search(r"As at (\d{4}-\d{2}-\d{2})", str(v))
                if m:
                    as_at = pd.Timestamp(m.group(1))
                    break
        except Exception:
            pass
        if as_at is None:
            raise ValueError("auto prior book: rpf uncached and as-at not "
                             "found on the Audit tab")

        def _datedif_m(d1, d2):
            m = (d2.year - d1.year) * 12 + (d2.month - d1.month)
            return m - 1 if d2.day < d1.day else m

        def _band(months):
            return 0.2 if months < 6 else 0.4 if months < 12 \
                else 0.6 if months < 18 else 0.8
        L["rpf"] = L["off_risk_date"].map(
            lambda d: _band(_datedif_m(as_at, d)) if pd.notna(d) else 0.0)
    else:
        L["rpf"] = pd.to_numeric(L["rpf"], errors="coerce")
    L["debris_dr"] = L["orbit"].map(DEBRIS_DR).fillna(0.0)
    L["on_risk_flag"] = 1
    L["layer_key"] = (L["program_id"].astype(str) + "|"
                      + L["layer_id"].astype(str))
    L["layer_occ"] = L.groupby("layer_key")["layer_key"].transform("size")
    L["exposure_usd"] = L["per_sc"] * L["layer_occ"]
    L["xol_ceded"] = 0.0
    L["net_of_xol"] = L["net_of_qs"]
    L["excluded_reason"] = None
    L["uw_year"] = L["inception"].dt.year
    L["is_consortium"] = (L["s3123_qs"].fillna(0) > 0) | \
        (L["equity_usd"].fillna(0) > 0)
    return L.reset_index(drop=True)


def restore_addons(prior, cur_layers, p):
    """Recompute the prior book's S3123 QS + IG equity add-ons from the engine's
    own formulas when they arrived as zero (an auto prior book stores them as
    uncached formulas). Without this the FIHL exec-basis (combined) prior
    collapses to raw, so the Change-Narrative Δ compares current-with-add-ons
    against prior-without — not like-for-like.

    is_consortium is taken from the CURRENT book by spacecraft (the same
    satellite carries the same consortium status quarter to quarter), because the
    auto book's Per Layer tab does not carry the flag. Idempotent: if the prior
    already holds non-zero add-ons (cached values), it is returned unchanged.
    """
    if prior is None or not len(prior) or p is None:
        return prior
    have = 0.0
    for c in ("s3123_qs", "equity_usd"):
        if c in prior.columns:
            have += float(pd.to_numeric(prior[c], errors="coerce").fillna(0).sum())
    if have > 0:
        return prior

    def _sid(v):
        n = pd.to_numeric(v, errors="coerce")
        return None if pd.isna(n) else int(n)

    # consortium status per spacecraft, from the current book
    cons = {}
    if cur_layers is not None and "spacecraft_id" in getattr(cur_layers, "columns", []):
        flag = cur_layers.get("is_consortium")
        if flag is None:
            flag = ((pd.to_numeric(cur_layers.get("s3123_qs"), errors="coerce").fillna(0) > 0)
                    | (pd.to_numeric(cur_layers.get("equity_usd"), errors="coerce").fillna(0) > 0))
        for sid, ic in zip(cur_layers["spacecraft_id"], flag):
            k = _sid(sid)
            if k is not None:
                cons[k] = bool(ic) or cons.get(k, False)

    q = p.s3123_qs
    cs = q.get("consortium_start")
    cons_start = pd.Timestamp(cs).date() if cs else None
    date_from = pd.Timestamp(q["date_from"]).date()
    if cons_start:
        date_from = max(date_from, cons_start)
    date_to = pd.Timestamp(q["date_to"]).date()
    excl = {int(x) for x in (q.get("excluded", []) or [])}
    cession = float(q["cession"])

    prior = prior.copy()
    s3, eq = [], []
    for _, row in prior.iterrows():
        inc = row.get("inception")
        inc = pd.Timestamp(inc).date() if pd.notna(inc) else None
        sid = _sid(row.get("spacecraft_id"))
        is_cons = cons.get(sid, False)
        if inc is not None and cons_start and inc < cons_start:
            is_cons = False
        px = float(row.get("per_sc") or 0.0)
        fac = p.s3123_factor(inc) if inc else 0.0
        eq.append(px * p.equity_pct(inc, is_cons) * fac if inc else 0.0)
        elig = (is_cons and inc is not None and sid not in excl
                and date_from <= inc <= date_to)
        s3.append(px * cession * fac if elig else 0.0)
    prior["s3123_qs"] = s3
    prior["equity_usd"] = eq
    return prior


# --------------------------------------------------------------------------- #
# changes dict — exactly the schema excel_report._changes / _exposure_bridge /
# waterfalls.build_movement_waterfalls consume
# --------------------------------------------------------------------------- #
def _nid(s):
    """'344664' from int 344664, float 344664.0, or '344664.0' alike —
    FIX: SQL-ingested ids arrive as floats; raw astype(str) produced
    '344664.0' vs the prior's '344664', matching NOTHING (the
    '153 new / 139 dropped / 0 moved' symptom)."""
    return pd.to_numeric(s, errors="coerce").astype("Int64").astype(str)


def _key(df):
    return (_nid(df["program_id"]) + "|" + _nid(df["layer_id"])
            + "|" + _nid(df["spacecraft_id"]))


def build_changes(cur_layers: pd.DataFrame, cur_grid: pd.DataFrame,
                  prior_layers: pd.DataFrame, prior_grid: pd.DataFrame,
                  prior_as_at: str) -> dict:
    cur = cur_layers.copy(); pri = prior_layers.copy()
    cur["_k"] = _key(cur); pri["_k"] = _key(pri)
    ck, pk = set(cur["_k"]), set(pri["_k"])

    new = cur[cur["_k"].isin(ck - pk)].copy()
    dropped = pri[pri["_k"].isin(pk - ck)].copy()
    common = sorted(ck & pk)
    cm = cur.set_index("_k").loc[common, "per_sc"]
    pm = pri.set_index("_k").loc[common, "per_sc"]
    # multi-spacecraft duplicate keys are impossible (key includes spacecraft),
    # but guard against accidental duplicates in either book
    cm = cm.groupby(level=0).sum(); pm = pm.groupby(level=0).sum()
    move = cm - pm
    moved = move[move.abs() > 1]

    # Renewal-gap flag: dropped layers whose spacecraft is absent from the
    # current book entirely (no continuing or renewed layer). A 1-July annual
    # tranche that expires at quarter-close with its renewals not yet bound /
    # entered in the source surfaces here — the trough looks like a loss until
    # the renewals come in. Distinguishes a genuine non-renewal from a timing gap.
    cur_scid = (set(cur["spacecraft_id"].dropna().astype(str))
                if "spacecraft_id" in cur.columns else set())
    if "spacecraft_id" in dropped.columns and cur_scid:
        gaps = dropped[~dropped["spacecraft_id"].astype(str).isin(cur_scid)].copy()
    else:
        gaps = dropped.iloc[0:0].copy()

    layers = {
        "summary": {
            "new_layers": int(len(new)),
            "dropped_layers": int(len(dropped)),
            "moved_layers": int(len(moved)),
            "new_exposure": float(new["per_sc"].sum()),
            "dropped_exposure": float(dropped["per_sc"].sum()),
            "net_move": float(move.sum()),
            "renewal_gap_layers": int(len(gaps)),
            "renewal_gap_exposure": float(gaps["per_sc"].sum()) if len(gaps) else 0.0,
        },
        "new": new, "dropped": dropped, "renewal_gaps": gaps,
    }

    # scenario movement — join the two grids on entity × scenario.
    # exec-basis columns mirror the MANUAL Executive Summary matrix convention:
    # FIHL = combined (incl S3123 QS + equity), FIBL = internal receiver.
    def _with_exec(g, gname, ename):
        g = g.copy()
        exec_g = g["gross"].astype(float).copy()
        net_v = g["net"].astype(float).copy()
        if "gross_incl_addons" in g.columns:
            m = g["entity"].eq("FIHL")
            exec_g[m] = g.loc[m, "gross_incl_addons"].astype(float)
        if "gross_receiver" in g.columns:
            m = g["entity"].eq("FIBL")
            exec_g[m] = g.loc[m, "gross_receiver"].astype(float)
        # FIBL headline net is the RECEIVER net (matches its receiver gross); its
        # direct-book "net" is ~0 for internally-ceded birds and would misreport
        # the receiver row (e.g. Max Risk SPAINSAT NG-2 net would show 0).
        if "net_receiver" in g.columns:
            m = g["entity"].eq("FIBL")
            net_v[m] = g.loc[m, "net_receiver"].astype(float)
        out = g[["entity", "scenario"]].copy()
        out[gname] = g["gross"].astype(float).values
        out[ename] = net_v.values
        out[gname + "_exec"] = exec_g.values
        return out
    c = _with_exec(cur_grid, "cur_gross", "cur_net")
    p = _with_exec(prior_grid, "prior_gross", "prior_net")
    sc = c.merge(p, on=["entity", "scenario"], how="outer")
    sc["d_gross"] = sc["cur_gross"] - sc["prior_gross"]
    sc["d_net"] = sc["cur_net"] - sc["prior_net"]
    sc["d_gross_pct"] = np.where(sc["prior_gross"].fillna(0) != 0,
                                 sc["d_gross"] / sc["prior_gross"], np.nan)

    return {"prior_as_at": prior_as_at, "scenarios": sc,
            "layers": layers, "prior_layers": prior_layers}
