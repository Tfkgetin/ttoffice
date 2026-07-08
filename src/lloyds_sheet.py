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


def write_lloyds_rds_summary(wb, data, as_at="", prior=None):
    """data = {'rds': df|None, 'space_weather': df|None} from
    ingest.load_lloyds_rds_summary. `prior` (optional) = a prior-quarter
    {RDS_Name: RDS_Value} dict for the Change-vs-Last-Quarter column."""
    if "Lloyds RDS Summary" in wb.sheetnames:
        del wb["Lloyds RDS Summary"]
    ws = wb.create_sheet("Lloyds RDS Summary")
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGH", [3, 60, 22, 26, 18, 16, 16, 20]):
        ws.column_dimensions[col].width = w

    r = 2
    ws.cell(r, 2, "Lloyd's RDS Summary — Syndicate 3123").font = _f(16, True, NAVY)
    r += 1
    ws.cell(r, 2, f"Space · S3123 · as at {as_at} · from "
            f"rds.vw_SpaceRDS_All_Lloyds_RDS / vw_SpaceRDS_SpaceWeather_Lloyds"
            ).font = _f(10, False, SOFT)
    r += 2

    rds = (data or {}).get("rds")
    sw = (data or {}).get("space_weather")

    # ---------------- Block 1: the four headline RDS ---------------------- #
    ws.cell(r, 2, "RDS — realistic disaster scenarios").font = _f(12, True, NAVY)
    r += 1
    if rds is None or not len(rds):
        r = _banner(ws, r, "vw_SpaceRDS_All_Lloyds_RDS returned no rows — the "
                    "view is not configured or not reachable from this login. "
                    "Set s3123_rds.lloyds_summary.rds_view and re-run on SQL.")
    else:
        c_name = _pick(rds, "RDS_Name") or rds.columns[0]
        c_val = _pick(rds, "RDS_Value") or rds.columns[1]
        c_brk = _pick(rds, "Breaches_Risk_Appetite", "BreachesRiskAppetite")
        has_prior = bool(prior)
        cols = ["RDS", "RDS Value (USD)", "Breaches risk appetite?"] \
            + (["Last quarter (USD)", "Change (USD)"] if has_prior else [])
        widths = [60, 22, 20] + ([18, 16] if has_prior else [])
        r = _header(ws, r, cols, widths)
        for _, row in rds.iterrows():
            name = row[c_name]
            ws.cell(r, 2, str(name)).font = _f(10, True, INK)
            ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
            vc = ws.cell(r, 3, _money(row[c_val]))
            vc.font = _f(10, False, INK)
            vc.alignment = Alignment(horizontal="right")
            breach = _is_yes(row[c_brk]) if c_brk else False
            bc = ws.cell(r, 4, "Yes" if breach else "No")
            bc.font = _f(10, True, RED if breach else GREEN)
            bc.alignment = Alignment(horizontal="right")
            if breach:
                for c in range(2, 5):
                    ws.cell(r, c).fill = _fill(REDFILL)
            if has_prior:
                pv = prior.get(str(name))
                pc = ws.cell(r, 5, _money(pv)); pc.font = _f(10, False, SOFT)
                pc.alignment = Alignment(horizontal="right")
                try:
                    d = float(row[c_val]) - float(pv)
                except (TypeError, ValueError):
                    d = None
                dc = ws.cell(r, 6, _money(d) if d is not None else "-")
                dc.alignment = Alignment(horizontal="right")
                dc.font = _f(10, False, GREEN if (d or 0) < 0 else RED if (d or 0) > 0 else SOFT)
            last_c = 6 if has_prior else 4
            for c in range(2, last_c + 1):
                ws.cell(r, c).border = Border(bottom=thin)
            r += 1
    r += 2

    # ---------- Block 2: Space Weather — design deficiency by bus type ---- #
    ws.cell(r, 2, "Space Weather — design deficiency by satellite bus type"
            ).font = _f(12, True, NAVY)
    r += 1
    ws.cell(r, 2, "Top-4 gross exposures on the largest bus-type group drive the "
            "Design Deficiency RDS above.").font = _f(9, False, SOFT)
    r += 1
    if sw is None or not len(sw):
        r = _banner(ws, r, "vw_SpaceRDS_SpaceWeather_Lloyds returned no rows — "
                    "set s3123_rds.lloyds_summary.space_weather_view and re-run "
                    "on SQL.")
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
    ws.cell(r, 2, "Source: the syndicate's own Lloyd's RDS views (same database "
            "as the IG book). Methodology — Generic Defect: 50% of the largest "
            "manufacturer's fleet (RPF-adjusted); Proton Flare: 5% of all GEO; "
            "Space Debris – Group 1: 100% of the worst LEO altitude group "
            "(RPF-adjusted); Space Weather – Design Deficiency: top-4 gross on "
            "the largest bus-type group.").font = _f(9, False, SOFT)
    ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    ws.row_dimensions[r].height = 42
    return ws
