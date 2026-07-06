"""Renewal-policy classification.

Ground truth for "did an expiring layer renew?" is the underwriter's own
forward pointer in the PBI layer snapshot (J.Pbi.Layers_t):

    Renewed To Program Id   – successor programme (0 / NULL = none)
    Renewed To UW Status    – successor's status (Bound / Quoted-Awaiting FOT /
                              NTU / NBI / Declined …)

This replaces the old spacecraft-id heuristic (which mis-called fleet refreshes
and could not tell a genuine non-renewal from a renewal not yet bound). The
classifier is pure/testable; the SQL that supplies the columns is optional, so
everything degrades to the heuristic when the pointers are absent.
"""
from __future__ import annotations
import pandas as pd

# state -> (display label, severity). Severity drives the workbook colour:
#   gap      = renewed & bound but MISSING from our source  → action (strong red)
#   lost     = NTU / Declined                               → genuine loss (red)
#   progress = renewal quoted / in pipeline, not yet bound  → manual-incl. candidate (amber)
#   ok       = renewed and already in our book              → fine (green)
#   neutral  = no pointer at all                            → candidate lapse (grey)
STATE_META = {
    "bound_missing": ("Bound — missing from source", "gap"),
    "lost":          ("Not renewed (NTU/Declined)",  "lost"),
    "in_progress":   ("In progress (unbound)",       "progress"),
    "bound_captured":("Renewed (in book)",           "ok"),
    "no_pointer":    ("No renewal pointer",          "neutral"),
}

_LOST = ("ntu", "declined", "dead", "not taken", "lapsed")
_BOUND = ("bound", "signed", "live", "inforce", "in force")


def classify(renewed_to, status, in_book: bool) -> str:
    """Classify one expiring layer from its forward pointer + successor status.

    renewed_to : successor programme id (None/NA/0 → no renewal recorded)
    status     : successor UW status string (may be None)
    in_book    : is the successor programme already in the current RDS book?
    """
    none_like = renewed_to is None or str(renewed_to) in ("0", "0.0", "<NA>", "nan", "")
    if not none_like:
        try:
            none_like = bool(pd.isna(renewed_to))
        except (TypeError, ValueError):
            none_like = False
    if none_like:
        return "no_pointer"
    s = str(status or "").strip().lower()
    if any(t in s for t in _LOST):
        return "lost"
    if any(t in s for t in _BOUND):
        return "bound_captured" if in_book else "bound_missing"
    # a target exists but it is quoted / awaiting FOT / NBI / pipeline / unknown
    return "in_progress"


def label(state: str) -> str:
    return STATE_META.get(state, ("—", "neutral"))[0]


def severity(state: str) -> str:
    return STATE_META.get(state, ("—", "neutral"))[1]


def annotate(df: pd.DataFrame, current_pids) -> pd.DataFrame:
    """Add `renewal_state` / `renewal_label` columns to a frame of expiring
    layers. No-op (returns df unchanged) if the pointer column is absent — the
    caller then falls back to the spacecraft heuristic."""
    if df is None or not len(df) or "renewed_to_program_id" not in df.columns:
        return df
    pids = {str(int(p)) for p in current_pids
            if p is not None and not pd.isna(p)} if current_pids is not None else set()

    def _pid(v):
        if v is None or pd.isna(v):
            return None
        try:
            return str(int(v))
        except (TypeError, ValueError):
            return str(v)

    out = df.copy()
    st = out.get("renewed_to_status")
    states = []
    for i in range(len(out)):
        rt = out["renewed_to_program_id"].iloc[i]
        stat = st.iloc[i] if st is not None else None
        in_book = _pid(rt) in pids
        states.append(classify(rt, stat, in_book))
    out["renewal_state"] = states
    out["renewal_label"] = [label(s) for s in states]
    out["renewal_severity"] = [severity(s) for s in states]
    return out


def summarize(states) -> dict:
    """Count states (accepts a Series/list). Returns {state: n} for present states."""
    s = pd.Series(list(states))
    return {k: int(v) for k, v in s.value_counts().items()}


def _layer_key(pid, lid):
    def _n(v):
        try:
            return str(int(float(v)))
        except (TypeError, ValueError):
            return str(v)
    return f"{_n(pid)}_{_n(lid)}"


def load_watch(params):
    """Load the renewal-watch set (expiring space layers + forward pointers) from
    the PBI snapshot. Optional: returns a DataFrame keyed by `layer_key`, or None
    when SQL / pyodbc / the `renewal_watch_query_file` config are unavailable —
    the caller then falls back to the spacecraft heuristic. Never raises."""
    ing = getattr(params, "ingest", {}) or {}
    if ing.get("source") != "sql":
        return None
    sql_cfg = dict(ing.get("sql") or {})
    qf = sql_cfg.get("renewal_watch_query_file")
    if not qf:
        return None
    try:
        import pyodbc  # noqa: F401
        from pathlib import Path
        from . import ingest as _ing
        driver = sql_cfg.get("driver") or _ing._pick_driver()
        conn = (f"DRIVER={{{driver}}};SERVER={sql_cfg['server']};"
                f"DATABASE={sql_cfg['database']};Trusted_Connection=yes;")
        if "18" in driver:
            conn += "TrustServerCertificate=yes;"
        cn = pyodbc.connect(conn, timeout=30)
        w = pd.read_sql(Path(qf).read_text(), cn)
        cn.close()
    except Exception as e:  # noqa: BLE001 — optional feature, never break the run
        print(f"      renewal-watch: skipped ({type(e).__name__}: {e})")
        return None
    w = w.rename(columns={c: _ing._norm(c) for c in w.columns})
    ren = {"programid": "program_id", "layerid": "layer_id",
           "renewedtoprogramid": "renewed_to_program_id",
           "renewedfromprogramid": "renewed_from_program_id",
           "renewedtouwstatus": "renewed_to_status",
           "renewedtoinception": "renewed_to_inception"}
    w = w.rename(columns={k: v for k, v in ren.items() if k in w.columns})
    if "program_id" not in w.columns or "layer_id" not in w.columns:
        return None
    for col in ("renewed_to_program_id", "renewed_from_program_id"):
        if col in w.columns:
            w[col] = pd.to_numeric(w[col], errors="coerce").replace(0, pd.NA).astype("Int64")
    w["layer_key"] = [_layer_key(p, l) for p, l in zip(w["program_id"], w["layer_id"])]
    return w


def attach_watch(frame, watch, current_pids):
    """Merge renewal pointers from `watch` onto a frame of expiring/dropped
    layers (by layer_key) and classify. Returns the frame unchanged if either
    input is missing or the frame has no layer_key — safe to call always."""
    if frame is None or not len(frame) or watch is None or not len(watch):
        return frame
    f = frame.copy()
    if "layer_key" not in f.columns:
        if {"program_id", "layer_id"} <= set(f.columns):
            f["layer_key"] = [_layer_key(p, l)
                              for p, l in zip(f["program_id"], f["layer_id"])]
        else:
            return frame
    cols = ["layer_key", "renewed_to_program_id", "renewed_to_status"]
    cols = [c for c in cols if c in watch.columns]
    f = f.merge(watch[cols].drop_duplicates("layer_key"), on="layer_key",
                how="left", suffixes=("", "_w"))
    return annotate(f, current_pids)
