"""Scenario-level netting waterfalls — layer-exact, mirroring the workbook.

Each scenario defines a per-layer multiplier m_i on per-S/C exposure.
Every cession stage then aggregates as sum(m_i * stage_i), and IGR XoL is
applied to the scenario-aggregate net of QS (Netting Waterfall tab logic).

Verified 2026Q1 (recon): this module, fed the manual Input Data, reproduces the
manual Summary to the dollar on every FUL/FIID/FIHL cell (incl. FIID Space
Weather XoL 9,928,305 and all FIHL blocks). All auto-vs-manual differences were
data-side or presentational — see the bridge workbook.
"""
from __future__ import annotations
import pandas as pd


def scenario_multipliers(df: pd.DataFrame, p, entity: str | None = None) -> dict:
    """Per-layer multiplier series for each scenario (entity-specific for SW/MR)."""
    flag = df["on_risk_flag"] == 1
    geo = df["orbit"].eq("GEO-GSO") & flag
    geomeo = df["orbit"].isin(["GEO-GSO", "MEO"]) & flag
    all_orbits = df["orbit"].isin(["GEO-GSO", "MEO", "LEO"]) & flag

    m = {
        "Proton Flare": geo * p.scenarios["proton_flare"]["insured_loss"],
        "Generic Defect": geomeo * p.scenarios["generic_defect"]["insured_loss"] * df["rpf"],
        "Space Debris": all_orbits * df["debris_dr"] * p.scenarios["space_debris"]["insured_loss"],
    }

    ent_mask = df["entity"].eq(entity) if entity else pd.Series(True, index=df.index)

    # Space Weather: worst single bus manufacturer — ALL orbits (Bus Manufacturer
    # pivot sums the mfr's whole fleet). Ranked by net-of-QS for FUL/FIID
    # (J−K−L−M / Q−R−S−T in the pivot) and net-of-Ext-QS (BF) for FIHL.
    # NOTE: the metric inconsistency (net_of_qs per entity vs net_ext for group)
    # deliberately mirrors the manual workbook's own formulas; both picks agree
    # in 2026Q1 (Airbus Toulouse). Do not "fix" without agreeing a convention.
    sw_scope = flag & ent_mask
    rank_metric = (df["net_ext"] if entity is None
                   else df["per_sc"] - df["ext_qs"] - df["igr_qs_ceded"])
    grp = rank_metric[sw_scope].groupby(df.loc[sw_scope, "bus_manufacturer"]).sum()
    if len(grp):
        worst_mfr = grp.idxmax()
        m["Space Weather"] = (flag & df["bus_manufacturer"].eq(worst_mfr)).astype(float)
        m["_sw_detail"] = worst_mfr
    # Max Risk: largest spacecraft by per-S/C gross for this perspective
    mr_scope = flag & ent_mask
    sc = df[mr_scope].groupby("spacecraft_name")["per_sc"].sum()
    if len(sc):
        biggest = sc.idxmax()
        m["Max Risk"] = df["spacecraft_name"].eq(biggest).astype(float) * flag
        m["_mr_detail"] = biggest
    return m


def fibl_maxrisk_multiplier(df: pd.DataFrame):
    """Max Risk multiplier for FIBL: the single on-risk spacecraft FIBL RECEIVES
    the most on — IGR QS ceded up from FUL/FIID + S3123 QS + IG equity + any
    direct FIBL line. FIBL is the internal RI receiver, so its worst single-
    satellite loss is the internally-ceded bird, NOT the group's largest-GROSS
    bird (which the entity=None Max Risk pick hands it — e.g. SXM-10, a FUL bird,
    of which FIBL only receives a thin IGR slice). Ranking AND value then use the
    SAME bird, so the figure is a coherent single-satellite loss — every
    component is that one bird's own cession.

    (Basis B, UW-confirmed 2026Q2: reproduces SPAINSAT NG-2 receiver 21,402,848
    on the 2026Q1 manual Input Data. The filed manual's 22,946,346 instead
    stitched the max IGR of SPAINSAT NG-1 onto NG-2's S3123/equity — not a single
    loss; superseded.)

    Returns (multiplier_series, spacecraft_name), or (None, None) if empty.
    """
    flag = df["on_risk_flag"] == 1
    equity = df.get("equity_usd", 0.0)
    received = (df["igr_qs_ceded"].fillna(0)
                + df["s3123_qs"].fillna(0)
                + (equity.fillna(0) if hasattr(equity, "fillna") else equity)
                + df["per_sc"] * df["entity"].eq("FIBL"))
    sc = (received * flag).groupby(df["spacecraft_name"]).sum()
    if not len(sc) or sc.max() <= 0:
        return None, None
    biggest = sc.idxmax()
    return df["spacecraft_name"].eq(biggest).astype(float) * flag, biggest


def _net(df: pd.DataFrame, m: pd.Series, entity: str, p) -> dict:
    """Layer-exact waterfall for one entity × scenario multiplier."""
    e = df["entity"].eq(entity)
    w = m * e
    gross = (w * df["per_sc"]).sum()
    ext_qs = (w * df["ext_qs"]).sum()
    other = (w * df["other_ext_ri"]).sum()
    net_ext = (w * df["net_ext"]).sum()
    qs_ceded = (w * df["igr_qs_ceded"]).sum()
    net_of_qs = (w * df["net_of_qs"]).sum()
    xol = p.xol_recovery(entity, net_of_qs)
    return {
        "gross": gross, "ext_qs": ext_qs, "other_ext_ri": other,
        "net_of_ext_qs": gross - ext_qs, "igr_qs_ceded": qs_ceded,
        "net_of_qs": net_of_qs, "xol_ceded": xol, "net": net_of_qs - xol,
    }


def _fihl_fibl(df: pd.DataFrame, m: pd.Series, p) -> dict:
    """FIHL (group) and FIBL (internal RI receiver) for one scenario multiplier.

    Verified to the dollar against the manual Netting Waterfall (combined table):
      FIHL = all-GEO gross + S3123 QS to IG + IG Equity Share
             (group/consolidated; gross and net carry the same IG add-ons)
      FIBL = QS ceded (from FUL+FIID) + XoL ceded + direct FIBL + S3123 + Equity
             (the internal receiver; gross ≈ net, only the external QS on its
             direct book differs)

    The scenario multiplier m already encodes orbit gate × InsLoss × RPF, so we
    just weight the per-layer columns the engine produced (per_sc, s3123_qs,
    equity_usd, igr_qs_ceded, xol_ceded).

    DISPLAY (recon 2026Q1, C1 — DECIDED: match the manual convention):
      * FIHL "gross"/"net"  = EX add-ons (all-entity gross / net-of-ext-QS),
        exactly the manual Summary row. The folded group view lives in
        "gross_incl_addons"/"net_incl_addons" and renders on the
        "S3123 & Equity (IG)" tab.
      * FIBL "gross"/"net"  = DIRECT book only, exactly the manual Summary row.
        The receiver view (QS in + XoL in + direct + add-ons) lives in
        "gross_receiver"/"net_receiver" and renders on the same tab.
      This also puts the QoQ Changes matrix on the manual basis, so movement
      reconciles with manual-vs-manual up to data corrections.
    """
    w = m
    # IG add-ons (apply to the whole on-risk population the scenario touches)
    s3123 = (w * df["s3123_qs"]).sum()
    equity = (w * df.get("equity_usd", 0)).sum()

    # --- FIHL: group — headline EX add-ons (manual Summary convention) ---
    all_gross = (w * df["per_sc"]).sum()
    all_ext_qs = (w * df["ext_qs"]).sum()
    all_net_ext = (w * df["net_ext"]).sum()
    fihl = {
        "entity": "FIHL", "gross": all_gross, "ext_qs": all_ext_qs,
        "net_of_ext_qs": all_gross - all_ext_qs,
        "igr_qs_ceded": 0.0, "net_of_qs": all_net_ext,
        "xol_ceded": 0.0, "net": all_net_ext,
        "s3123_qs": s3123, "equity_usd": equity,
        "gross_incl_addons": all_gross + s3123 + equity,
        "net_incl_addons": all_net_ext + s3123 + equity,
    }

    # --- FIBL: headline = DIRECT book (manual Summary convention);
    #     receiver view retained in *_receiver keys ---
    # FIBL receives QS + XoL ceded from FUL and FIID. XoL is computed at the
    # scenario-aggregate level (not per-layer), so recompute each operating
    # entity's XoL recovery on its aggregate net-of-QS, exactly as _net does.
    qs_ceded_in = (w * df["igr_qs_ceded"]).sum()      # QS ceded from FUL+FIID
    xol_ceded_in = 0.0
    for op in ("FUL", "FIID"):
        e = df["entity"].eq(op)
        op_net_of_qs = (w * df["net_of_qs"] * e).sum()
        xol_ceded_in += p.xol_recovery(op, op_net_of_qs)
    direct_fibl = (w * df["per_sc"] * df["entity"].eq("FIBL")).sum()
    direct_fibl_net = (w * df["net_ext"] * df["entity"].eq("FIBL")).sum()
    fibl = {
        "entity": "FIBL", "gross": direct_fibl,
        "ext_qs": direct_fibl - direct_fibl_net,
        "net_of_ext_qs": direct_fibl_net,
        "igr_qs_ceded": 0.0, "net_of_qs": direct_fibl_net,
        "xol_ceded": 0.0, "net": direct_fibl_net,
        "s3123_qs": s3123, "equity_usd": equity,
        "qs_received": qs_ceded_in, "xol_received": xol_ceded_in,
        "direct_fibl": direct_fibl,
        # receiver view (manual Netting Waterfall combined table, verified):
        "gross_receiver": qs_ceded_in + xol_ceded_in + direct_fibl
                          + s3123 + equity,
        "net_receiver": qs_ceded_in + xol_ceded_in + direct_fibl_net
                        + s3123 + equity,
    }
    return {"FIHL": fihl, "FIBL": fibl}


def entity_scenario_netting(df: pd.DataFrame, p) -> pd.DataFrame:
    """Netting Waterfall tab: per entity × scenario, layer-exact.

    Builds all four entities — FUL & FIID (operating, full cede-down cascade),
    plus FIHL (group) and FIBL (internal RI receiver) via the verified combined
    formula. Scenario multipliers are taken from the group (entity=None) pass so
    FIHL/FIBL see the whole population; FUL/FIID use their own perspective for
    Space Weather / Max Risk ranking.
    """
    rows = []
    # FUL & FIID — operating entities, full cascade, own SW/MR perspective
    for entity in ["FUL", "FIID"]:
        ms = scenario_multipliers(df, p, entity=entity)
        details = {"Space Weather": ms.pop("_sw_detail", None),
                   "Max Risk": ms.pop("_mr_detail", None)}
        for scen, m in ms.items():
            wf = _net(df, m, entity, p)
            rows.append({"entity": entity, "scenario": scen,
                         "detail": details.get(scen), **wf})

    # FIHL & FIBL — group + receiver, off the group-level multipliers
    msg = scenario_multipliers(df, p, entity=None)
    gdetails = {"Space Weather": msg.pop("_sw_detail", None),
                "Max Risk": msg.pop("_mr_detail", None)}
    fibl_mr_m, fibl_mr_sc = fibl_maxrisk_multiplier(df)
    for scen, m in msg.items():
        ff = _fihl_fibl(df, m, p)
        rows.append({"scenario": scen, "detail": gdetails.get(scen),
                     **ff["FIHL"]})
        # FIBL Max Risk uses its OWN receiver-ranked bird (SPAINSAT NG-2), not
        # the group's largest-gross pick (SXM-10) the other scenarios share.
        # See fibl_maxrisk_multiplier: value and label are the same single bird.
        if scen == "Max Risk" and fibl_mr_m is not None:
            ff_fibl = _fihl_fibl(df, fibl_mr_m, p)
            rows.append({"scenario": scen, "detail": fibl_mr_sc,
                         **ff_fibl["FIBL"]})
        else:
            rows.append({"scenario": scen, "detail": gdetails.get(scen),
                         **ff["FIBL"]})
    return pd.DataFrame(rows)


def summary_grid(df: pd.DataFrame, p) -> pd.DataFrame:
    """All four entities × all scenarios, gross + net, layer-exact."""
    return entity_scenario_netting(df, p)


def fihl_selection_contribs(df: pd.DataFrame, p) -> pd.DataFrame:
    """Per-layer FIHL contribution to the SELECTION scenarios (Space Weather,
    Max Risk), with the group selection baked in — so the 'S3123 & Equity (IG)'
    tab can render the FIHL ex-add-on gross as a live =SUM(Per Layer) for these
    scenarios too (the additive ones already have pf/gd/sd_fihl).

      sw_fihl = m_SpaceWeather × per_sc  (per_sc over the worst-manufacturer
                on-risk layers, all orbits)
      mr_fihl = m_MaxRisk × per_sc       (per_sc over the largest-spacecraft
                on-risk layers)

    m is the GROUP (entity=None) multiplier, exactly as _fihl_fibl uses it, so
    SUM(sw_fihl) reproduces the grid's FIHL Space-Weather gross to the cent and
    SUMPRODUCT(sw_fihl, s3123_qs/per_sc) reproduces the S3123-QS add-back
    ((m × s3123_qs).sum()). Same for Max Risk."""
    m = scenario_multipliers(df, p, entity=None)
    out = pd.DataFrame(index=df.index)
    if "Space Weather" in m:
        out["sw_fihl"] = (m["Space Weather"] * df["per_sc"]).astype(float)
    if "Max Risk" in m:
        out["mr_fihl"] = (m["Max Risk"] * df["per_sc"]).astype(float)
    return out
