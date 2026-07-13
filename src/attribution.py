"""Scenario Δ attribution — decompose each entity × scenario GROSS move
(Q-on-Q) into named drivers, so the input-template explanation is specific, not
general.

Every spacecraft lands in exactly one bucket, so the four buckets sum to the
grid Δ to the cent:

    Δ = New business  −  Run-off  ±  Renewals (reprice)  ±  Continuing (reval/RPF/selection)

Current per-layer scenario contributions come off the Per Layer backbone
(pf/gd/sd/sw/mr_{fihl,ful,fiid}); the PRIOR contributions are recomputed by
running the SAME engine contribution functions on the frozen prior per-layer
snapshot, so both sides are engine-exact and the decomposition ties.
"""
from __future__ import annotations
import pandas as pd

from . import scenarios as _scen, netting as _net, renewals as _rnw

SCEN_SLUG = {"Proton Flare": "pf", "Generic Defect": "gd", "Space Debris": "sd",
             "Space Weather": "sw", "Max Risk": "mr"}
ENT_SLUG = {"FIHL": "fihl", "FUL": "ful", "FIID": "fiid"}


def _prep(df, params):
    df = df.copy()
    if "debris_dr" not in df.columns:
        dbo = (params.scenarios.get("space_debris", {}) or {}).get("damage_by_orbit", {})
        df["debris_dr"] = df.get("orbit", "").map(lambda o: float(dbo.get(o, 0.0)))
    for col, dflt in (("on_risk_flag", 1), ("rpf", 1.0), ("igr_qs_rate", 0.0),
                      ("xol_ceded", 0.0), ("other_ext_ri", 0.0)):
        if col not in df.columns:
            df[col] = dflt
    return df


def _with_contribs(df, params):
    """Attach pf/gd/sd_{fihl,ful,fiid} + sw/mr_{fihl,ful,fiid} via the engine."""
    p = _prep(df, params)
    add = pd.concat([_scen.proton_flare(p, params), _scen.generic_defect(p, params),
                     _scen.space_debris(p, params),
                     _net.selection_contribs(p, params)], axis=1)
    return p.join(add)


def compute(per_layer, prior_layers, changes, params):
    """{(entity, scenario): {prior, cur, delta, buckets{}, drivers{}}} or None."""
    if prior_layers is None or not len(prior_layers):
        return None
    cur = per_layer
    # current contributions already on Per Layer; add selection ones if absent
    need = [f"{s}_{e}" for s in SCEN_SLUG.values() for e in ENT_SLUG.values()]
    if any(c not in cur.columns for c in need):
        cur = _with_contribs(cur, params)
    pri = _with_contribs(prior_layers, params)

    lyr = (changes or {}).get("layers", {}) or {}
    sp = _rnw.split_movement(lyr.get("new"), lyr.get("dropped"))
    ren_sid = set(sp.get("ren_sid") or set())

    cur = cur.assign(_sid=_rnw._idn_col(cur, "spacecraft_id"))
    pri = pri.assign(_sid=_rnw._idn_col(pri, "spacecraft_id"))
    name = (cur.groupby("_sid")["spacecraft_name"].first()
            .combine_first(pri.groupby("_sid")["spacecraft_name"].first()))

    out = {}
    for ent, es in ENT_SLUG.items():
        for scen, ss in SCEN_SLUG.items():
            col = f"{ss}_{es}"
            if col not in cur.columns or col not in pri.columns:
                continue
            cb = cur.groupby("_sid")[col].sum()
            pb = pri.groupby("_sid")[col].sum()
            cur_ids, pri_ids = set(cb.index), set(pb.index)
            buckets = {"new": 0.0, "renewals": 0.0, "continuing": 0.0, "lapsed": 0.0}
            drivers = {"new": [], "renewals": [], "continuing": [], "lapsed": []}
            for sid in cur_ids | pri_ids:
                c, pr = float(cb.get(sid, 0.0)), float(pb.get(sid, 0.0))
                if sid in cur_ids and sid not in pri_ids:
                    b, val = "new", c
                elif sid in pri_ids and sid not in cur_ids:
                    b, val = "lapsed", -pr
                elif sid in ren_sid:
                    b, val = "renewals", c - pr
                else:
                    b, val = "continuing", c - pr
                buckets[b] += val
                if abs(val) > 1.0:
                    drivers[b].append((str(name.get(sid, sid)), val))
            for b in drivers:
                drivers[b].sort(key=lambda x: -abs(x[1]))
            ct, pt = float(cur[col].sum()), float(pri[col].sum())
            out[(ent, scen)] = {"prior": pt, "cur": ct, "delta": ct - pt,
                                "buckets": buckets,
                                "drivers": {b: v[:5] for b, v in drivers.items()}}
    return out
