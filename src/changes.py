"""Q-on-Q changes: current run vs a prior persisted run."""
from __future__ import annotations
import pandas as pd

SCEN_ORDER = ["Proton Flare", "Space Weather", "Generic Defect",
              "Space Debris", "Max Risk"]


def layer_changes(cur: pd.DataFrame, prior: pd.DataFrame) -> dict:
    """New / dropped / moved layers between two runs (keyed on layer_key)."""
    cur = cur.copy(); prior = prior.copy()
    cur["layer_key"] = cur["layer_key"].astype(str)
    prior["layer_key"] = prior["layer_key"].astype(str)
    ck = set(cur["layer_key"]); pk = set(prior["layer_key"])

    new_keys = ck - pk
    dropped_keys = pk - ck
    common = ck & pk

    new = cur[cur["layer_key"].isin(new_keys)].copy()
    dropped = prior[prior["layer_key"].isin(dropped_keys)].copy()

    # exposure moves on layers present in both
    c = cur[cur["layer_key"].isin(common)].set_index("layer_key")["per_sc"]
    p = prior[prior["layer_key"].isin(common)].set_index("layer_key")["per_sc"]
    moved = pd.DataFrame({"prior": p, "current": c})
    moved["delta"] = moved["current"] - moved["prior"]
    moved = moved[moved["delta"].abs() > 1].copy()
    if len(moved):
        meta = cur.set_index("layer_key")[
            ["spacecraft_name", "entity", "bus_manufacturer"]]
        moved = moved.join(meta).sort_values("delta", key=abs, ascending=False)

    # Renewal-gap flag: dropped layers whose spacecraft is absent from the
    # current book entirely (no continuing/renewed layer) — candidates for a
    # renewal not yet bound. Separates a real non-renewal from a timing trough.
    cur_scid = (set(cur["spacecraft_id"].dropna().astype(str))
                if "spacecraft_id" in cur.columns else set())
    if "spacecraft_id" in dropped.columns and cur_scid:
        gaps = dropped[~dropped["spacecraft_id"].astype(str).isin(cur_scid)].copy()
    else:
        gaps = dropped.iloc[0:0].copy()

    return {
        "new": new, "dropped": dropped, "moved": moved.reset_index(),
        "renewal_gaps": gaps,
        "summary": {
            "new_layers": len(new), "dropped_layers": len(dropped),
            "moved_layers": len(moved),
            "new_exposure": float(new["per_sc"].sum()) if len(new) else 0.0,
            "dropped_exposure": float(dropped["per_sc"].sum()) if len(dropped) else 0.0,
            "net_move": float(moved["delta"].sum()) if len(moved) else 0.0,
            "renewal_gap_layers": len(gaps),
            "renewal_gap_exposure": float(gaps["per_sc"].sum()) if len(gaps) else 0.0,
        },
    }


def scenario_changes(cur_grid: pd.DataFrame, prior_grid: pd.DataFrame) -> pd.DataFrame:
    """Gross & net deltas per entity × scenario.

    Covers all four entities — FIHL (group), FUL & FIID (operating), and FIBL
    (internal RI receiver) — so the Changes matrix and the by-scenario net/gross
    tables show the complete entity picture.
    """
    def val(df, f):
        if len(df) and f in df and pd.notna(df.iloc[0][f]):
            return float(df.iloc[0][f])
        return None

    # Mirror prior_seed._with_exec so this fallback path shows the SAME headline
    # basis as the frozen-workbook path: FIHL gross = combined (incl S3123 QS +
    # equity), FIBL gross/net = internal receiver. Without this, FIBL would show
    # its ~0 direct-book net (blank prior net) and no exec-gross column.
    def _exec_gross(df, entity):
        if entity == "FIHL":
            v = val(df, "gross_incl_addons")
            if v is not None:
                return v
        if entity == "FIBL":
            v = val(df, "gross_receiver")
            if v is not None:
                return v
        return val(df, "gross")

    def _head_net(df, entity):
        # FIBL headline net is the receiver net (matches its receiver gross);
        # its direct-book net is ~0. FIHL headline net is COMBINED (incl S3123
        # QS + IG equity) to match its combined gross — without this FIHL net
        # would show ex-add-on against a combined gross (the basis mismatch that
        # made prior/current net wrong). FUL/FIID keep their standard net.
        if entity == "FIBL":
            v = val(df, "net_receiver")
            if v is not None:
                return v
        if entity == "FIHL":
            v = val(df, "net_incl_addons")
            if v is not None:
                return v
        return val(df, "net")

    rows = []
    for entity in ["FIHL", "FUL", "FIBL", "FIID"]:
        for scen in SCEN_ORDER:
            cg = cur_grid[(cur_grid["entity"] == entity) &
                          (cur_grid["scenario"] == scen)]
            pg = prior_grid[(prior_grid["entity"] == entity) &
                            (prior_grid["scenario"] == scen)]
            if not len(cg) and not len(pg):
                continue
            cgr, pgr = val(cg, "gross"), val(pg, "gross")
            cnet, pnet = _head_net(cg, entity), _head_net(pg, entity)
            ceg, peg = _exec_gross(cg, entity), _exec_gross(pg, entity)
            rows.append({
                "entity": entity, "scenario": scen,
                "cur_gross": cgr, "prior_gross": pgr,
                "cur_gross_exec": ceg, "prior_gross_exec": peg,
                "d_gross": (cgr - pgr) if None not in (cgr, pgr) else None,
                "d_gross_pct": ((cgr - pgr) / pgr) if (pgr) else None,
                "cur_net": cnet, "prior_net": pnet,
                "d_net": (cnet - pnet) if None not in (cnet, pnet) else None,
            })
    return pd.DataFrame(rows)


def s3123_changes(cur_s3: "pd.DataFrame | None",
                  prior_s3: "pd.DataFrame | None") -> "pd.DataFrame | None":
    """Q-on-Q for the S3123 (Lloyd's) grid — same structure as scenario_changes
    but only entity='S3123'."""
    if cur_s3 is None or prior_s3 is None:
        return None
    rows = []
    for scen in SCEN_ORDER:
        cg = cur_s3[cur_s3["scenario"] == scen]
        pg = prior_s3[prior_s3["scenario"] == scen]
        if not len(cg) and not len(pg):
            continue
        def val(df, f):
            if len(df) and f in df.columns and pd.notna(df.iloc[0][f]):
                return float(df.iloc[0][f])
            return None
        cgr, pgr = val(cg, "gross"), val(pg, "gross")
        cnet, pnet = val(cg, "net"), val(pg, "net")
        rows.append({
            "entity": "S3123", "scenario": scen,
            "cur_gross": cgr, "prior_gross": pgr,
            "d_gross": (cgr - pgr) if None not in (cgr, pgr) else None,
            "d_gross_pct": ((cgr - pgr) / pgr) if pgr else None,
            "cur_net": cnet, "prior_net": pnet,
            "d_net": (cnet - pnet) if None not in (cnet, pnet) else None,
        })
    return pd.DataFrame(rows) if rows else None


def compute(cur_layers, cur_grid, prior, cur_s3=None):
    """prior = (layers_df, grid_df, as_at[, s3123_grid]) from persist.load_prior,
    or None."""
    if prior is None:
        return None
    # handle both 3-tuple (old) and 4-tuple (new, with S3123 grid)
    if len(prior) == 4:
        p_layers, p_grid, p_asat, p_s3 = prior
    else:
        p_layers, p_grid, p_asat = prior
        p_s3 = None
    result = {
        "prior_as_at": p_asat,
        "layers": layer_changes(cur_layers, p_layers),
        "scenarios": scenario_changes(cur_grid, p_grid),
    }
    s3diff = s3123_changes(cur_s3, p_s3)
    if s3diff is not None:
        result["s3123"] = s3diff
    return result
