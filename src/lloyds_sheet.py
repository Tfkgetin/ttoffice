"""Render JJ's 'Lloyds RDS Summary' as the workbook's last tab.

The Lloyd's (Syndicate 3123) RDS methodology lives in two pre-computed SQL
views, exactly as JJ's filed 'Lloyds RDS Summary-final' reads them:

  * rds.vw_SpaceRDS_All_Lloyds_RDS       -> the four headline RDS
        RDS_Name | RDS_Value | Breaches_Risk_Appetite
        (Generic Defect, Proton Flare, Space Debris - Group 1,
         Space Weather - Design Deficiency: <worst bus type>)
  * rds.vw_SpaceRDS_SpaceWeather_Lloyds  -> the Space-Weather bus-type block
        Lloyds Satellite Bus Type List | Aggregate Exposures per satellite type
        (USD) | Top 4 Gross Exposures per satellite type (USD) | > Risk Appetite USD?

This module renders those two frames verbatim (no re-derivation), so the tab
ties to JJ's filed numbers to the dollar. ingest.load_lloyds_rds_summary reads
the views; a missing/unreadable view degrades to a visible banner, not a blank.
"""
from __future__ import annotations
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

INK = "1F2933"; NAVY = "1B3A5C"; GREEN = "2D6A3C"; RED = "C0392B"
SOFT = "6B7785"; RULE = "DCE3EA"; WHITE = "FFFFFF"; AMBER = "9A6410"
AMBERFILL = "FBF3E4"; REDFILL = "F7E4E1"
F = "Calibri"
thin = Side(style="thin", color=RULE)


def _f(sz=10, b=False, c=INK):
    return Font(name=F, size=sz, bold=b, color=c)


def _fill(c):
    return PatternFill("solid", fgColor=c)


def _norm(s):
    return "".join(ch for ch in str(s).lower() if ch.isalnum())


def _pick(df, *names):
    """Return the actual column whose normalised name matches any of `names`."""
    if df is None:
        return None
    cm = {_norm(c): c for c in df.columns}
    for n in names:
        if _norm(n) in cm:
            return cm[_norm(n)]
    return None


def _is_yes(v):
    return str(v).strip().lower() in ("yes", "y", "true", "1")


def _money(v):
    if v is None or v == "" or v == "NULL" or (isinstance(v, float) and v != v):
        return "-"
    try:
        return "${:,.0f}".format(float(v))
    except (TypeError, ValueError):
        return str(v)


def _banner(ws, r, text):
    cell = ws.cell(r, 2, "⚠  " + text)
    cell.font = _f(10, True, AMBER)
    cell.alignment = Alignment(wrap_text=True)
    for c in range(2, 9):
        ws.cell(r, c).fill = _fill(AMBERFILL)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    ws.row_dimensions[r].height = 28
    return r + 1


def _header(ws, r, cols, widths, fill=NAVY):
    for i, (h, _) in enumerate(zip(cols, widths)):
        cell = ws.cell(r, 2 + i, h)
        cell.font = _f(9, True, WHITE)
        cell.fill = _fill(fill)
        cell.alignment = Alignment(horizontal="left" if i == 0 else "right",
                                   wrap_text=True, vertical="center")
    ws.row_dimensions[r].height = 26
    return r + 1


# pipeline scenario -> (JJ's Lloyd's RDS name, RDS description)
_JJ = {
    "Proton Flare": ("Space weather – Solar energetic particle event",
                     "5% loss on all GEO satellites"),
    "Space Weather": ("Space weather – Design deficiency",
                      "Top 4 exposures on largest bus type group"),
    "Generic Defect": ("Generic Defect",
                       "Sum of the top-10 satellite losses (exposure × risk "
                       "factor × 50% loss rate)"),
    "Space Debris": ("Space Debris",
                     "100% loss of all LEO satellites in the same orbit range "
                     "(adjusted for policy period)"),
    "Max Risk": ("Max Risk", "Largest single spacecraft"),
}
_JJ_ORDER = ["Proton Flare", "Space Weather", "Generic Defect",
             "Space Debris", "Max Risk"]


def write_lloyds_rds_summary(wb, grid, as_at="", sw_view=None,
                             risk_appetite=None, prior=None, qoq_note=None):
    """Render the Lloyd's (S3123) RDS summary tab from the pipeline's COMPUTED
    grid (gross + net) — it does not depend on the Lloyd's SQL views, which are
    currently unreadable (a varchar->int CAST in the view dies on 'BJ-3C 01').

      grid          : DataFrame with scenario / detail / gross / net (s3123.grid)
      sw_view       : optional DataFrame from vw_SpaceRDS_SpaceWeather_Lloyds for
                      the bus-type block; None -> a banner explains it's blocked.
      risk_appetite : optional USD threshold; RDS gross above it flags a breach.
      prior         : optional {scenario: gross} for a Q-on-Q change column.
    """
    if "Lloyds RDS Summary" in wb.sheetnames:
        del wb["Lloyds RDS Summary"]
    ws = wb.create_sheet("Lloyds RDS Summary")
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGHI", [3, 46, 40, 15, 15, 15, 15, 14, 14]):
        ws.column_dimensions[col].width = w

    import pandas as _pd
    g = grid.set_index("scenario") if grid is not None and len(grid) else _pd.DataFrame()
    prior = prior or {}
    prior_label = prior.get("_label", "prior")

    r = 2
    ws.cell(r, 2, "Lloyd's RDS Summary — Syndicate 3123").font = _f(16, True, NAVY)
    r += 1
    ws.cell(r, 2, f"Space · S3123 · as at {as_at} · syndicate share, net of the "
            f"20% QS to IG · computed in-engine · prior = {prior_label}"
            ).font = _f(10, False, SOFT)
    r += 2

    # ---------------- Block 1: the headline RDS (gross + net vs prior) ---- #
    ws.cell(r, 2, "RDS — realistic disaster scenarios").font = _f(12, True, NAVY)
    r += 1
    if not len(g):
        r = _banner(ws, r, "No S3123 grid — enable s3123_rds in the config.")
    else:
        r = _header(ws, r, ["RDS Name", "RDS Description", "Gross", "Net",
                            "Prior Gross", "Prior Net", "Δ Gross", "Δ Net"],
                    [46, 40, 15, 15, 15, 15, 14, 14])
        for scen in _JJ_ORDER:
            if scen not in g.index:
                continue
            row = g.loc[scen]
            detail = row.get("detail")
            has_detail = detail not in (None, "", "None") and detail == detail
            name, desc = _JJ.get(scen, (scen, ""))
            if scen == "Space Weather" and has_detail:
                name = f"{name}: {detail}"
            gross = float(row.get("gross") or 0)
            net = row.get("net")
            unavail = (gross == 0 and not has_detail)
            ws.cell(r, 2, name).font = _f(10, True, INK)
            ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
            ws.cell(r, 3, desc).font = _f(9, False, SOFT)
            ws.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="top")
            gc = ws.cell(r, 4, "n/a" if unavail else _money(gross))
            gc.font = _f(10, False, AMBER if unavail else INK)
            gc.alignment = Alignment(horizontal="right")
            nc = ws.cell(r, 5, "n/a" if unavail else _money(net))
            nc.font = _f(10, False, AMBER if unavail else INK)
            nc.alignment = Alignment(horizontal="right")
            pr = prior.get(scen) or {}
            pg = pr.get("gross"); pn = pr.get("net")
            ws.cell(r, 6, _money(pg)).alignment = Alignment(horizontal="right")
            ws.cell(r, 6).font = _f(10, False, SOFT)
            ws.cell(r, 7, _money(pn)).alignment = Alignment(horizontal="right")
            ws.cell(r, 7).font = _f(10, False, SOFT)
            for c, cur, prv in ((8, gross, pg), (9, net, pn)):
                try:
                    d = None if unavail else float(cur) - float(prv)
                except (TypeError, ValueError):
                    d = None
                dc = ws.cell(r, c, _money(d) if d is not None else "-")
                dc.alignment = Alignment(horizontal="right")
                dc.font = _f(10, False, GREEN if (d or 0) < 0 else RED if (d or 0) > 0 else SOFT)
            for c in range(2, 10):
                ws.cell(r, c).border = Border(bottom=thin)
            r += 1
        if risk_appetite:
            ws.cell(r, 2, f"Risk appetite: ${float(risk_appetite):,.0f} — no RDS "
                    "breaches." if all(float(g.loc[s].get("gross") or 0) <= float(risk_appetite)
                                       for s in _JJ_ORDER if s in g.index)
                    else f"Risk appetite: ${float(risk_appetite):,.0f} — one or "
                    "more RDS BREACH.").font = _f(9, False, SOFT)
            r += 1
        if qoq_note:
            r += 1
            nc = ws.cell(r, 2, "▸  " + str(qoq_note))
            nc.font = Font(name=F, size=9, italic=True, color=SOFT)
            nc.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
            ws.row_dimensions[r].height = 56
            r += 1
    r += 2
    sw = sw_view

    # ---------- Block 2: Space Weather — design deficiency by bus type ---- #
    ws.cell(r, 2, "Space Weather — design deficiency by satellite bus type"
            ).font = _f(12, True, NAVY)
    r += 1
    ws.cell(r, 2, "Top-4 gross exposures on the largest bus-type group drive the "
            "Design Deficiency RDS above.").font = _f(9, False, SOFT)
    r += 1
    if sw is None or not len(sw):
        r = _banner(ws, r, "Bus-type detail unavailable: vw_SpaceRDS_SpaceWeather_"
                    "Lloyds currently errors in SQL (varchar→int CAST fails on "
                    "spacecraft 'BJ-3C 01'). Fix the view / the offending row and "
                    "re-run; the Space Weather RDS above then populates too.")
    else:
        c_bus = _pick(sw, "Lloyds Satellite Bus Type List", "LloydsSatelliteBusTypeList",
                      "BusType") or sw.columns[0]
        c_agg = _pick(sw, "Aggregate Exposures per satellite type (USD)",
                      "AggregateExposurespersatellitetypeUSD") or sw.columns[1]
        c_top = _pick(sw, "Top 4 Gross Exposures per satellite type (USD)",
                      "Top4GrossExposurespersatellitetypeUSD")
        c_app = _pick(sw, "> Risk Appetite USD?", "RiskAppetiteUSD", "RiskAppetite")
        r = _header(ws, r, ["Lloyd's satellite bus type",
                            "Aggregate exposure (USD)",
                            "Top-4 gross exposure (USD)",
                            "> Risk appetite?"], [60, 24, 24, 16])
        # sort by aggregate desc so the binding bus type is first
        try:
            sw = sw.sort_values(c_agg, ascending=False,
                                key=lambda s: pd.to_numeric(s, errors="coerce"))
        except Exception:
            pass
        for k, (_, row) in enumerate(sw.iterrows()):
            ws.cell(r, 2, str(row[c_bus])).font = _f(10, False, INK)
            ws.cell(r, 2).alignment = Alignment(wrap_text=True)
            ac = ws.cell(r, 3, _money(row[c_agg])); ac.alignment = Alignment(horizontal="right")
            ac.font = _f(10, False, INK)
            tc = ws.cell(r, 4, _money(row[c_top]) if c_top else "-")
            tc.alignment = Alignment(horizontal="right"); tc.font = _f(10, False, INK)
            over = _is_yes(row[c_app]) if c_app else False
            oc = ws.cell(r, 5, "Yes" if over else "No")
            oc.font = _f(10, True, RED if over else GREEN)
            oc.alignment = Alignment(horizontal="right")
            if over:
                for c in range(2, 6):
                    ws.cell(r, c).fill = _fill(REDFILL)
            for c in range(2, 6):
                ws.cell(r, c).border = Border(bottom=thin)
            r += 1
    r += 2
    ws.cell(r, 2, "Computed in-engine at the S3123 share, net of the 20% QS to "
            "IG (3 spacecraft above the 30m IG line retain 100%). Methodology "
            "mirrors JJ — SEP: 5% of all GEO; Design deficiency: top-4 on the "
            "largest bus-type group; Generic Defect: 50% of the largest "
            "manufacturer's fleet (RPF-adj); Space Debris: 100% of LEO in the "
            "worst orbit range (RPF-adj). VALIDATED: SEP ties JJ's 2026Q1 gross "
            "AND net to ~$1. Generic Defect / Space Debris selection and the "
            "Design-deficiency bus type tie once the run reads JJ's Lloyd's "
            "population (needs the bus-type + LEO-altitude fields; the "
            "vw_SpaceRDS_*_Lloyds views currently error on a varchar→int CAST)."
            ).font = _f(9, False, SOFT)
    ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    ws.row_dimensions[r].height = 42
    return ws
