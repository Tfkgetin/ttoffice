"""Scenario loss calculators — per-layer losses for the 5 RDS scenarios."""
from __future__ import annotations
import pandas as pd


def proton_flare(df: pd.DataFrame, p) -> pd.DataFrame:
    """GEO-GSO fleet × insured-loss %. No time decay (workbook BY-CD)."""
    s = p.scenarios["proton_flare"]
    in_scope = df["orbit"].isin(s["orbits"]) & (df["on_risk_flag"] == 1)
    loss = df["per_sc"] * s["insured_loss"]

    out = pd.DataFrame(index=df.index)
    out["pf_fihl"] = (in_scope * loss).astype(float)
    out["pf_ful"] = (in_scope & (df["entity"] == "FUL")) * loss
    out["pf_fiid"] = (in_scope & (df["entity"] == "FIID")) * loss
    out["pf_fibl_direct"] = (in_scope & (df["entity"] == "FIBL")) * loss
    # FIBL Ceded = (per_sc × QS rate − per-layer XoL) × loss%
    out["pf_fibl_ceded"] = in_scope * (
        df["per_sc"] * df["igr_qs_rate"] - df["xol_ceded"]) * s["insured_loss"]
    return out


def generic_defect(df: pd.DataFrame, p) -> pd.DataFrame:
    """(GEO|MEO) × 50% × RPF (workbook CR-CW)."""
    s = p.scenarios["generic_defect"]
    in_scope = df["orbit"].isin(s["orbits"]) & (df["on_risk_flag"] == 1)
    loss = df["per_sc"] * s["insured_loss"] * df["rpf"]

    out = pd.DataFrame(index=df.index)
    out["gd_fihl"] = in_scope * loss
    out["gd_ful"] = (in_scope & (df["entity"] == "FUL")) * loss
    out["gd_fiid"] = (in_scope & (df["entity"] == "FIID")) * loss
    out["gd_fibl_direct"] = (in_scope & (df["entity"] == "FIBL")) * loss
    out["gd_fibl_ceded"] = in_scope * (
        df["per_sc"] * df["igr_qs_rate"] - df["xol_ceded"]) * s["insured_loss"] * df["rpf"]
    return out


def space_debris(df: pd.DataFrame, p) -> pd.DataFrame:
    """Orbit DR × 100% (workbook Summary D8: GEO+MEO+LEO).
    NOTE: the Input Data CY-DD per-layer block excludes LEO; reconciliation
    against wb_sd_* therefore compares the GEO+MEO subset only."""
    s = p.scenarios["space_debris"]
    in_scope = df["orbit"].isin(s["orbits"]) & (df["on_risk_flag"] == 1)
    loss = df["per_sc"] * s["insured_loss"] * df["debris_dr"]

    out = pd.DataFrame(index=df.index)
    out["sd_fihl"] = in_scope * loss
    out["sd_ful"] = (in_scope & (df["entity"] == "FUL")) * loss
    out["sd_fiid"] = (in_scope & (df["entity"] == "FIID")) * loss
    out["sd_fibl_direct"] = (in_scope & (df["entity"] == "FIBL")) * loss
    out["sd_fibl_ceded"] = in_scope * (
        df["per_sc"] * df["igr_qs_rate"] - df["xol_ceded"]) * s["insured_loss"] * df["debris_dr"]
    return out


def space_weather_groups(df: pd.DataFrame) -> pd.DataFrame:
    """Manufacturer-group totals per entity, ALL on-risk orbits (raw gross
    per_sc, before any equity/netting). The worst manufacturer is picked
    downstream. Matches the manual Bus Manufacturer tab basis."""
    sc = df[df["on_risk_flag"] == 1]
    g = sc.groupby("bus_manufacturer").agg(
        sw_fihl=("per_sc", "sum"),
        sw_fibl_ceded=("total_fibl_ceded", "sum"),
    )
    for ent, col in [("FUL", "sw_ful"), ("FIID", "sw_fiid"), ("FIBL", "sw_fibl_direct")]:
        g[col] = sc[sc["entity"] == ent].groupby("bus_manufacturer")["per_sc"].sum()
    return g.fillna(0.0).reset_index()


def max_risk_layers(df: pd.DataFrame) -> pd.DataFrame:
    """Per-layer totals (multi-spacecraft rows summed back) — workbook BT/BU.
    Largest layer per perspective is picked downstream."""
    g = df.groupby("layer_key").agg(
        program_id=("program_id", "first"),
        layer_id=("layer_id", "first"),
        spacecraft_name=("spacecraft_name", "first"),
        entity=("entity", "first"),
        maxrisk_gross=("net_ext", "sum"),     # BT: SUMIFS(BF)
        maxrisk_net=("net_of_xol", "sum"),    # BU: SUMIFS(BL)
        gross_exposure=("per_sc", "sum"),
    )
    return g.reset_index()


def run_scenarios(df: pd.DataFrame, p):
    per_layer = pd.concat(
        [df, proton_flare(df, p), generic_defect(df, p), space_debris(df, p)], axis=1)
    sw = space_weather_groups(per_layer)
    mr = max_risk_layers(per_layer)
    return per_layer, sw, mr
