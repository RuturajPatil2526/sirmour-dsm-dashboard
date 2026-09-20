"""
excel_export.py
================
Builds the downloadable DSM penalty report as a single, formula-driven Excel
sheet that mirrors the plant's standard "Schedule vs Meter + Penalty"
template:

    Row 1        Title
    Row 2        Note
    Row 4        Column headers (Block, Time, AI Schedule (MW), Actual (MW),
                 Deviation (MW), Deviation % (Capacity), Penalty (Rs),
                 Scheduled at)
    Rows 5..N    One row per matched block. Deviation, Deviation % and
                 Penalty are LIVE EXCEL FORMULAS (not pre-computed numbers)
                 that reference the DSM slab parameter table at the bottom
                 of the sheet, so the workbook is fully auditable/editable
                 in Excel.
    (blank)
    DAY SUMMARY section    -- accuracy & penalty roll-up, formula driven
    (blank)
    DSM SLAB PARAMETERS    -- installed capacity, block-energy factor and
                               the slab table the Penalty formulas reference

Every formula cell is written with XlsxWriter's write_formula(), which lets
us store a CACHED VALUE alongside the formula string (the value already
computed in Python by calculator.py). This means the numbers display
correctly immediately -- in Excel/LibreOffice/Google Sheets, in a quick
preview pane, or when read back with pandas -- without requiring the
consuming application to run its own formula engine. Opening the file in a
real spreadsheet app still shows (and recalculates) the live formula, so the
report stays fully auditable.

Conditional formatting highlights blocks with >0.5 MW deviation in red,
<=0.5 MW in blue, and any block that actually incurred a penalty in red.
"""

from __future__ import annotations

import io
from typing import List

import numpy as np
import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell, xl_col_to_name

from config import PlantConfig
from calculator import compute_daily_status

# All 6 charts are built as NATIVE Excel chart objects (not images), so this
# module has no dependency on kaleido / a headless browser at all -- charts
# are just formulas + xlsxwriter chart definitions, which always work the
# same way on every machine. Native charts are also fully interactive
# (hover tooltips, zoom, edit) and stay live if the data changes.
DARK_GREEN = "#1B4332"
WHITE = "#FFFFFF"
GRAY_TEXT = "#595959"
RED = "#C00000"
BLUE = "#1F4E9C"
PINK_FILL = "#FCE4E4"
BORDER_GRAY = "#BFBFBF"
INPUT_BLUE = "#0000FF"
COLOR_ENERCAST = "#2563EB"  # blue -- Enercast comparison series on charts (matches graphs.py)
COLOR_STEP1 = "#DB2777"     # pink/magenta -- Step 1 comparison series on charts (matches graphs.py)
COLOR_STEP2 = "#9333EA"     # purple -- Step 2 comparison series on charts (matches graphs.py)

HEADERS = [
    "Block", "Time", "AI Schedule (MW)", "Actual (MW)", "Deviation (MW)",
    "Deviation % (Capacity)", "Penalty (Rs)", "Scheduled at",
]
COL_WIDTHS = [8, 15, 17, 14, 15, 20, 13, 15]

# Optional columns I-L, added only when an Enercast (or other third-party
# forecast) file was uploaded. Enercast is comparison-only -- it never
# affects the AI Schedule's own Deviation/Penalty columns (C-G) above.
ENERCAST_HEADERS = [
    "Enercast (MW)", "Enercast Deviation (MW)",
    "Enercast Deviation % (Capacity)", "Enercast Penalty (Rs)",
]
ENERCAST_COL_WIDTHS = [13, 18, 20, 15]
ENERCAST_NOTE = (
    " Enercast is shown alongside for comparison only — our forecast is "
    "graded against actual meter data, never against Enercast."
)

# Optional column groups added only when the AI Schedule file used the
# multi-step format ("Step 1 Meter Base Forecast MW" -> "Step 2 Weather
# Adjustment MW" -> "Step 3 Plant Profile Adjustment MW"). Whichever stage
# is LAST (highest-numbered) present is the official schedule used
# throughout the report above (columns C-G); every earlier stage that's
# also present gets its own 4 columns here, penalised separately,
# comparison-only, so the value each later stage adds can be seen directly.
# Appended right after the Enercast columns (if present), in stage order,
# so both optional blocks stay contiguous and every other hardcoded column
# reference in this module (C/D/E/F/G, I/J/K/L for Enercast) is unaffected.
STEP_COLOR = {1: COLOR_STEP1, 2: COLOR_STEP2}


def _stage_headers(n: int) -> List[str]:
    return [
        f"Step {n} Forecast (MW)", f"Step {n} Deviation (MW)",
        f"Step {n} Deviation % (Capacity)", f"Step {n} Penalty (Rs)",
    ]


STAGE_COL_WIDTHS = [16, 18, 20, 15]
STAGE_STEP_NAMES = {1: "Step 1 (meter-base forecast)", 2: "Step 2 (weather-adjusted forecast)"}

# Always-present trailing columns: PPA Amount (display-only reference figure)
# and Status (Calculated / Pending). Appended LAST (after Enercast, if
# present) so every existing hardcoded column-letter reference (C/D/E/F/G,
# I/J/K/L) used throughout this module stays valid unchanged.
PPA_STATUS_HEADERS = ["PPA Amount (Rs)", "Status"]
PPA_STATUS_COL_WIDTHS = [15, 13]
PENDING_NOTE = (
    " A block with a missing AI Schedule and/or Meter value is marked "
    "'Pending' and never treated as zero — its Deviation/Penalty cells are "
    "left blank and it is excluded from every total until real data arrives."
)


def _fmt_date(block_df: pd.DataFrame) -> str:
    dates = sorted({d for d in block_df["Date"].dropna().unique()})
    if not dates:
        return "N/A"
    if len(dates) == 1:
        return str(dates[0])
    return f"{dates[0]} to {dates[-1]}"


def build_excel_report(block_df: pd.DataFrame, summary: dict, plant: PlantConfig, figures: dict = None) -> bytes:
    # `figures` is accepted for backward compatibility with older callers but
    # is no longer used -- charts are built natively from worksheet formulas,
    # not rendered from the Plotly figures / kaleido.
    output = io.BytesIO()
    wb = __import__("xlsxwriter").Workbook(output, {"in_memory": True})
    ws = wb.add_worksheet("Schedule vs Meter + Penalty")
    ws_graphs = wb.add_worksheet("Graphs")
    SHEET = "Schedule vs Meter + Penalty"

    n = len(block_df)
    has_enercast = "Enercast_MW" in block_df.columns
    has_status = "Status" in block_df.columns
    # Every comparison "stage" beyond the official schedule (Step 1 and/or
    # Step 2, whichever are present as their own MW columns -- see
    # utils.parse_energy_file / calculator.evaluate_step1/evaluate_step2)
    # gets its own contiguous 4-column block, in stage order, right after
    # Enercast (if present). Built generically so 1 or 2 active stages share
    # one code path below instead of a copy-pasted block per stage.
    active_stages = [n_ for n_ in (1, 2) if f"Step{n_}_MW" in block_df.columns]
    headers = HEADERS + (ENERCAST_HEADERS if has_enercast else [])
    col_widths = COL_WIDTHS + (ENERCAST_COL_WIDTHS if has_enercast else [])
    stage_col0 = {}
    stage_cols = {}  # n -> {"mw":, "dev":, "pct":, "pen":} column letters
    next_col_0 = len(HEADERS) + (len(ENERCAST_HEADERS) if has_enercast else 0)
    for n_ in active_stages:
        stage_col0[n_] = next_col_0
        stage_cols[n_] = {
            "mw": xl_col_to_name(next_col_0), "dev": xl_col_to_name(next_col_0 + 1),
            "pct": xl_col_to_name(next_col_0 + 2), "pen": xl_col_to_name(next_col_0 + 3),
        }
        headers += _stage_headers(n_)
        col_widths += STAGE_COL_WIDTHS
        next_col_0 += 4
    headers += PPA_STATUS_HEADERS
    col_widths += PPA_STATUS_COL_WIDTHS
    last_col = len(headers) - 1
    # PPA Amount / Status are always appended LAST (after Enercast and any
    # stage columns) so every existing hardcoded column-letter reference
    # below (C/D/E/F/G, I/J/K/L for Enercast) keeps pointing at exactly what
    # it always has.
    ppa_col_0 = next_col_0
    status_col_0 = ppa_col_0 + 1
    ppa_col = xl_col_to_name(ppa_col_0)
    status_col = xl_col_to_name(status_col_0)
    header_row_0 = 3          # 0-indexed row for the column header row (Excel row 4)
    data_start_0 = 4          # 0-indexed first data row (Excel row 5)
    data_end_0 = data_start_0 + n - 1

    calc_mask = (block_df["Status"] == "Calculated") if has_status else pd.Series([True] * n, index=block_df.index)
    n_calc = int(calc_mask.sum())
    n_pending = n - n_calc

    def xl_row(zero_based: int) -> int:
        return zero_based + 1  # 1-indexed Excel row number for formula strings

    # ------------------------------------------------------------- Formats
    title_fmt = wb.add_format({"bold": True, "font_size": 12, "font_color": DARK_GREEN, "font_name": "Arial"})
    note_fmt = wb.add_format({"italic": True, "font_size": 9, "font_color": GRAY_TEXT, "font_name": "Arial"})
    header_fmt = wb.add_format({
        "bold": True, "font_size": 10, "font_color": WHITE, "bg_color": DARK_GREEN,
        "align": "center", "valign": "vcenter", "text_wrap": True, "border": 1,
        "border_color": BORDER_GRAY, "font_name": "Arial",
    })
    section_fmt = wb.add_format({"bold": True, "font_size": 12, "font_color": DARK_GREEN, "font_name": "Arial"})
    label_fmt = wb.add_format({"font_size": 10, "font_name": "Arial"})
    input_fmt = wb.add_format({"font_size": 10, "font_color": INPUT_BLUE, "font_name": "Arial"})
    total_label_fmt = wb.add_format({"bold": True, "font_size": 10, "font_color": RED, "font_name": "Arial"})
    total_value_fmt = wb.add_format({
        "bold": True, "font_size": 10, "font_color": RED, "bg_color": PINK_FILL,
        "num_format": "0.00", "font_name": "Arial",
    })
    cell_fmt = wb.add_format({"align": "center", "border": 1, "border_color": BORDER_GRAY})
    mw_fmt = wb.add_format({"align": "center", "border": 1, "border_color": BORDER_GRAY, "num_format": "0.000"})
    pct_fmt = wb.add_format({"align": "center", "border": 1, "border_color": BORDER_GRAY, "num_format": "+0.00;-0.00"})
    pen_fmt = wb.add_format({"align": "center", "border": 1, "border_color": BORDER_GRAY, "num_format": "0.00"})
    ppa_fmt = wb.add_format({"align": "center", "border": 1, "border_color": BORDER_GRAY, "num_format": "0.00"})
    slab_edge_fmt = wb.add_format({"num_format": "0.000"})
    red_font = wb.add_format({"font_color": RED})
    blue_font = wb.add_format({"font_color": BLUE})
    status_calc_fmt = wb.add_format({
        "align": "center", "border": 1, "border_color": BORDER_GRAY,
        "font_color": "#166534", "bold": True,
    })
    status_pending_fmt = wb.add_format({
        "align": "center", "border": 1, "border_color": BORDER_GRAY,
        "font_color": "#B45309", "bold": True, "bg_color": "#FEF3C7", "italic": True,
    })

    fmt_int = wb.add_format({"num_format": "0"})
    fmt_3dp = wb.add_format({"num_format": "0.000"})
    fmt_signed_3dp = wb.add_format({"num_format": "+0.000;-0.000"})
    fmt_2dp = wb.add_format({"num_format": "0.00"})

    # ---------------------------------------------------------------- Title
    report_date = _fmt_date(block_df)
    ws.merge_range(0, 0, 0, last_col,
                    f"{plant.name} — AI Schedule vs Actual with DSM Penalty — {report_date}", title_fmt)
    note_text = ("Auto-generated by the AI Schedule Evaluation Dashboard. "
                 "Penalty is graded against actual meter data using piecewise DSM slab logic.")
    if has_status and n_pending:
        note_text += PENDING_NOTE
    if has_enercast:
        note_text += ENERCAST_NOTE
    if active_stages:
        official_n = max(active_stages) + 1  # the stage right after the last comparison stage is official
        stage_list = " and ".join(STAGE_STEP_NAMES[n_] for n_ in active_stages)
        note_text += (
            f" {stage_list} {'is' if len(active_stages) == 1 else 'are'} shown alongside for "
            f"comparison — the official AI Schedule is Step {official_n}; earlier stage(s) are "
            f"penalised separately, purely for reference, against the same actual meter data."
        )
    ws.merge_range(1, 0, 1, last_col, note_text, note_fmt)

    # -------------------------------------------------------------- Header
    ws.set_row(header_row_0, 26.4)
    for i, h in enumerate(headers):
        ws.write(header_row_0, i, h, header_fmt)

    # ----------------------------------------------------------- Data rows
    has_generated_at = "Generated_At" in block_df.columns
    for offset, (_, row) in enumerate(block_df.iterrows()):
        r0 = data_start_0 + offset
        r = xl_row(r0)
        is_calc = (row.get("Status") == "Calculated") if has_status else True
        ws.write_number(r0, 0, int(row["Block"]), cell_fmt)
        ws.write_string(r0, 1, str(row["Time_Label"]), cell_fmt)
        if pd.notna(row["Scheduled_MW"]):
            ws.write_number(r0, 2, float(row["Scheduled_MW"]), mw_fmt)
        else:
            ws.write_blank(r0, 2, None, mw_fmt)
        if pd.notna(row["Actual_MW"]):
            ws.write_number(r0, 3, float(row["Actual_MW"]), mw_fmt)
        else:
            ws.write_blank(r0, 3, None, mw_fmt)
        # E (Deviation MW) is only a formula for a Calculated block -- for a
        # Pending block (schedule and/or meter missing) it is left a genuine
        # blank, NEVER a formula, so it can't silently read a missing input
        # as zero (D{r}-C{r} would otherwise treat a blank cell as 0).
        if is_calc:
            ws.write_formula(r0, 4, f"=D{r}-C{r}", mw_fmt, float(row["Deviation_MW"]))
        else:
            ws.write_blank(r0, 4, None, mw_fmt)
        # PPA Amount (Rs) -- reference only, computed whenever a schedule
        # value exists, independent of Status.
        if pd.notna(row.get("PPA_Amount")):
            ws.write_number(r0, ppa_col_0, float(row["PPA_Amount"]), ppa_fmt)
        else:
            ws.write_blank(r0, ppa_col_0, None, ppa_fmt)
        # Status -- Calculated / Pending
        status_val = row.get("Status", "Calculated") if has_status else "Calculated"
        ws.write_string(r0, status_col_0, str(status_val),
                         status_pending_fmt if status_val == "Pending" else status_calc_fmt)
        # F (Deviation %) and G (Penalty) filled after the slab table is laid
        # out below, since their formulas reference those cells.
        if has_generated_at:
            val = row.get("Generated_At")
            if pd.notna(val) and str(val).strip().lower() != "nan":
                ws.write_string(r0, 7, str(val), cell_fmt)
            else:
                ws.write_blank(r0, 7, None, cell_fmt)
        else:
            ws.write_blank(r0, 7, None, cell_fmt)

        # I: Enercast (MW) -- a plain literal value (or a genuine blank cell
        # if this block has no Enercast coverage). J/K/L (Enercast Deviation
        # / Deviation % / Penalty) are filled in the formula pass below,
        # once the slab table's row numbers are known -- and are left as
        # real blanks (not a formula) for blocks with no Enercast data, so
        # SUM/COUNT/SUMPRODUCT over the whole column behave correctly.
        if has_enercast:
            e_mw = row.get("Enercast_MW")
            if pd.notna(e_mw):
                ws.write_number(r0, 8, float(e_mw), mw_fmt)
            else:
                ws.write_blank(r0, 8, None, mw_fmt)

        # Each active comparison stage's MW column -- a plain literal value
        # (or a genuine blank cell if this block has no forecast for that
        # stage). The Deviation/Deviation %/Penalty columns that follow are
        # filled in the formula pass below, exactly like Enercast's J/K/L.
        for n_ in active_stages:
            stage_mw = row.get(f"Step{n_}_MW")
            col0 = stage_col0[n_]
            if pd.notna(stage_mw):
                ws.write_number(r0, col0, float(stage_mw), mw_fmt)
            else:
                ws.write_blank(r0, col0, None, mw_fmt)

    # --------------------------------------------------------- Day summary
    blank1_0 = data_end_0 + 1
    summary_header_0 = blank1_0 + 1
    ws.write(summary_header_0, 0, "DAY SUMMARY — ACCURACY AND DSM PENALTY", section_fmt)

    # dev/pen are CALCULATED-ONLY (no Pending rows) -- these numpy arrays feed
    # plain (non-NaN-aware) aggregations below, so they must never contain a
    # Pending block's NaN, exactly mirroring what the Excel formulas (COUNT /
    # SUMIF ... "<>" on blank D) already compute on the sheet itself.
    dev = block_df.loc[calc_mask, "Deviation_MW"].to_numpy() if n_calc else np.array([])
    pen = block_df.loc[calc_mask, "Total_Block_Penalty"].to_numpy() if n_calc else np.array([])
    ppa_total = float(block_df["PPA_Amount"].sum()) if "PPA_Amount" in block_df.columns else 0.0
    daily_status = compute_daily_status(block_df) if has_status else "Calculated"
    cap = plant.installed_capacity_mw
    s, e = xl_row(data_start_0), xl_row(data_end_0)
    status_range = f"{status_col}{s}:{status_col}{e}"
    ppa_range = f"{ppa_col}{s}:{ppa_col}{e}"

    daily_status_formula = (
        f'=IF(COUNTIF({status_range},"Calculated")=0,"Pending",'
        f'IF(COUNTIF({status_range},"Calculated")<{n},"Partially Calculated",'
        f'IF(SUM(G{s}:G{e})=0,"Zero Penalty","Calculated")))'
    )

    summary_defs = [
        ("Daily Status", daily_status_formula, None, daily_status),
        ("Blocks with a real meter reading", f"=COUNT(D{s}:D{e})", fmt_int, float(n_calc)),
        ("Blocks Pending (schedule/meter missing)", f'=COUNTIF({status_range},"Pending")', fmt_int,
         float(n_pending)),
        ("Total scheduled (MW, scored blocks)", f'=SUMIF(D{s}:D{e},"<>",C{s}:C{e})', fmt_3dp,
         float(block_df.loc[calc_mask, "Scheduled_MW"].sum()) if n_calc else 0.0),
        ("Total actual (MW)", f"=SUM(D{s}:D{e})", fmt_3dp,
         float(block_df.loc[calc_mask, "Actual_MW"].sum()) if n_calc else 0.0),
        ("Total deviation — actual minus scheduled (MW)", f"=SUM(E{s}:E{e})", fmt_signed_3dp,
         float(dev.sum()) if n_calc else 0.0),
        ("Mean absolute deviation (MW)", f"=SUMPRODUCT(ABS(E{s}:E{e}))/COUNT(E{s}:E{e})", fmt_3dp,
         float(np.abs(dev).mean()) if n_calc else 0.0),
        ("Max absolute deviation (MW)", f"=MAX(MAX(E{s}:E{e}),-MIN(E{s}:E{e}))", fmt_3dp,
         float(np.abs(dev).max()) if n_calc else 0.0),
        ("Blocks over 0.5 MW deviation (RED)", f'=SUMPRODUCT((ABS(E{s}:E{e})>0.5)*(D{s}:D{e}<>""))', fmt_int,
         float((np.abs(dev) > 0.5).sum()) if n_calc else 0.0),
        ("Blocks within 0.5 MW deviation (BLUE)", f'=SUMPRODUCT((ABS(E{s}:E{e})<=0.5)*(D{s}:D{e}<>""))', fmt_int,
         float((np.abs(dev) <= 0.5).sum()) if n_calc else 0.0),
        ("Mean absolute deviation (% of capacity)", None, fmt_2dp,
         float(np.abs(dev).mean() / cap * 100) if n_calc else 0.0),  # formula added after cap_row known
        ("Blocks that incurred a penalty", f'=COUNTIF(G{s}:G{e},">0")', fmt_int,
         float((pen > 0).sum()) if n_calc else 0.0),
        ("Worst single-block penalty (Rs)", f"=MAX(G{s}:G{e})", fmt_2dp, float(pen.max()) if n_calc else 0.0),
        ("Total PPA Amount (Rs, reference only)", f"=SUM({ppa_range})", fmt_2dp, ppa_total),
        ("TOTAL DSM PENALTY FOR THE DAY (Rs)", f"=SUM(G{s}:G{e})", fmt_2dp, float(pen.sum()) if n_calc else 0.0),
    ]
    summary_rows_0 = {}
    for i, (label, formula, numfmt, cached) in enumerate(summary_defs):
        r0 = summary_header_0 + 1 + i
        ws.write(r0, 0, label, label_fmt)
        summary_rows_0[label] = r0
        if formula is not None:
            ws.write_formula(r0, 2, formula, numfmt, cached)
        # else: filled in later once the capacity cell row is known

    total_row_0 = summary_rows_0["TOTAL DSM PENALTY FOR THE DAY (Rs)"]
    ws.write(total_row_0, 0, "TOTAL DSM PENALTY FOR THE DAY (Rs)", total_label_fmt)
    ws.write_formula(total_row_0, 2, f"=SUM(G{s}:G{e})", total_value_fmt, float(pen.sum()) if n_calc else 0.0)

    # --------------------------------------------- Enercast summary (optional)
    # A second, parallel Day-Summary block for Enercast -- reference only,
    # never feeds into the AI Schedule's own penalty above. Only built when
    # an Enercast file was uploaded.
    e_summary_rows_0 = {}
    e_mad_pct_row_0 = None
    if has_enercast:
        # Enercast is graded against the actual meter reading too -- a block
        # with Enercast data but a missing/Pending Actual has nothing to
        # compare against, so it's excluded here exactly like the sheet's
        # own SUMIF(I..,"<>",D..) formulas already exclude it.
        e_mask = block_df["Enercast_MW"].notna() & block_df["Actual_MW"].notna()
        e_n = int(e_mask.sum())
        e_dev = block_df.loc[e_mask, "Enercast_Deviation_MW"].to_numpy() if e_n else np.array([])
        e_pen = block_df.loc[e_mask, "Enercast_Penalty"].to_numpy() if e_n else np.array([])
        e_total_enercast_mw = float(block_df.loc[e_mask, "Enercast_MW"].sum()) if e_n else 0.0
        e_total_actual_mw = float(block_df.loc[e_mask, "Actual_MW"].sum()) if e_n else 0.0

        e_header_0 = total_row_0 + 2
        ws.write(e_header_0, 0, "ENERCAST SUMMARY — ACCURACY AND DSM PENALTY (reference only)", section_fmt)

        e_defs = [
            ("Blocks with Enercast + a real meter reading", f"=COUNT(I{s}:I{e})", fmt_int, float(e_n)),
            ("Total Enercast (MW, scored blocks)", f"=SUM(I{s}:I{e})", fmt_3dp, e_total_enercast_mw),
            ("Total actual (MW, scored blocks)", f'=SUMIF(I{s}:I{e},"<>",D{s}:D{e})', fmt_3dp, e_total_actual_mw),
            ("Total deviation — actual minus Enercast (MW)", f"=SUM(J{s}:J{e})", fmt_signed_3dp,
             float(e_dev.sum()) if e_n else 0.0),
            ("Mean absolute deviation (MW)", f"=SUMPRODUCT(ABS(J{s}:J{e}))/COUNT(J{s}:J{e})", fmt_3dp,
             float(np.abs(e_dev).mean()) if e_n else 0.0),
            ("Max absolute deviation (MW)", f"=MAX(MAX(J{s}:J{e}),-MIN(J{s}:J{e}))", fmt_3dp,
             float(np.abs(e_dev).max()) if e_n else 0.0),
            ("Blocks over 0.5 MW deviation (RED)", f'=SUMPRODUCT((ABS(J{s}:J{e})>0.5)*(I{s}:I{e}<>""))', fmt_int,
             float((np.abs(e_dev) > 0.5).sum()) if e_n else 0.0),
            ("Blocks within 0.5 MW deviation (BLUE)", f'=SUMPRODUCT((ABS(J{s}:J{e})<=0.5)*(I{s}:I{e}<>""))', fmt_int,
             float((np.abs(e_dev) <= 0.5).sum()) if e_n else 0.0),
            ("Mean absolute deviation (% of capacity)", None, fmt_2dp,
             float(np.abs(e_dev).mean() / cap * 100) if e_n else 0.0),  # formula added after cap_row known
            ("Blocks that incurred a penalty", f'=COUNTIF(L{s}:L{e},">0")', fmt_int,
             float((e_pen > 0).sum()) if e_n else 0.0),
            ("Worst single-block penalty (Rs)", f"=MAX(L{s}:L{e})", fmt_2dp, float(e_pen.max()) if e_n else 0.0),
            ("TOTAL ENERCAST DSM PENALTY FOR THE DAY (Rs)", f"=SUM(L{s}:L{e})", fmt_2dp,
             float(e_pen.sum()) if e_n else 0.0),
        ]
        for i, (label, formula, numfmt, cached) in enumerate(e_defs):
            r0 = e_header_0 + 1 + i
            ws.write(r0, 0, label, label_fmt)
            e_summary_rows_0[label] = r0
            if formula is not None:
                ws.write_formula(r0, 2, formula, numfmt, cached)

        e_mad_pct_row_0 = e_summary_rows_0["Mean absolute deviation (% of capacity)"]
        e_total_row_0 = e_summary_rows_0["TOTAL ENERCAST DSM PENALTY FOR THE DAY (Rs)"]
        ws.write(e_total_row_0, 0, "TOTAL ENERCAST DSM PENALTY FOR THE DAY (Rs)", total_label_fmt)
        ws.write_formula(e_total_row_0, 2, f"=SUM(L{s}:L{e})", total_value_fmt, float(e_pen.sum()) if e_n else 0.0)
        last_summary_row_0 = e_total_row_0
    else:
        last_summary_row_0 = total_row_0

    # ------------------------------------------- Stage summaries (optional)
    # A parallel Day-Summary block for each active comparison stage (Step 1
    # and/or Step 2, whichever aren't the official schedule) -- reference
    # only, never feeds into the official schedule's own penalty above.
    # Built generically so 1 or 2 active stages share one code path.
    stage_mad_pct_row_0 = {}
    for n_ in active_stages:
        mw_c, dev_c, pct_c, pen_c = (stage_cols[n_]["mw"], stage_cols[n_]["dev"],
                                       stage_cols[n_]["pct"], stage_cols[n_]["pen"])
        mw_col_name = f"Step{n_}_MW"
        dev_col_name = f"Step{n_}_Deviation_MW"
        pen_col_name = f"Step{n_}_Penalty"

        # Graded against the actual meter reading too -- a block with a
        # forecast for this stage but a missing/Pending Actual has nothing
        # to compare against, exactly like the Enercast block above.
        st_mask = block_df[mw_col_name].notna() & block_df["Actual_MW"].notna()
        st_n = int(st_mask.sum())
        st_dev = block_df.loc[st_mask, dev_col_name].to_numpy() if st_n else np.array([])
        st_pen = block_df.loc[st_mask, pen_col_name].to_numpy() if st_n else np.array([])
        st_total_mw = float(block_df.loc[st_mask, mw_col_name].sum()) if st_n else 0.0
        st_total_actual_mw = float(block_df.loc[st_mask, "Actual_MW"].sum()) if st_n else 0.0

        st_header_0 = last_summary_row_0 + 2
        ws.write(st_header_0, 0, f"STEP {n_} SUMMARY — DSM PENALTY (reference only)", section_fmt)

        st_defs = [
            (f"Blocks with Step {n_} + a real meter reading",
             f"=COUNT({mw_c}{s}:{mw_c}{e})", fmt_int, float(st_n)),
            (f"Total Step {n_} (MW, scored blocks)",
             f"=SUM({mw_c}{s}:{mw_c}{e})", fmt_3dp, st_total_mw),
            ("Total actual (MW, scored blocks)",
             f'=SUMIF({mw_c}{s}:{mw_c}{e},"<>",D{s}:D{e})', fmt_3dp, st_total_actual_mw),
            (f"Total deviation — actual minus Step {n_} (MW)",
             f"=SUM({dev_c}{s}:{dev_c}{e})", fmt_signed_3dp,
             float(st_dev.sum()) if st_n else 0.0),
            ("Mean absolute deviation (MW)",
             f"=SUMPRODUCT(ABS({dev_c}{s}:{dev_c}{e}))/COUNT({dev_c}{s}:{dev_c}{e})", fmt_3dp,
             float(np.abs(st_dev).mean()) if st_n else 0.0),
            ("Max absolute deviation (MW)",
             f"=MAX(MAX({dev_c}{s}:{dev_c}{e}),-MIN({dev_c}{s}:{dev_c}{e}))", fmt_3dp,
             float(np.abs(st_dev).max()) if st_n else 0.0),
            ("Blocks over 0.5 MW deviation (RED)",
             f'=SUMPRODUCT((ABS({dev_c}{s}:{dev_c}{e})>0.5)*({mw_c}{s}:{mw_c}{e}<>""))', fmt_int,
             float((np.abs(st_dev) > 0.5).sum()) if st_n else 0.0),
            ("Blocks within 0.5 MW deviation (BLUE)",
             f'=SUMPRODUCT((ABS({dev_c}{s}:{dev_c}{e})<=0.5)*({mw_c}{s}:{mw_c}{e}<>""))', fmt_int,
             float((np.abs(st_dev) <= 0.5).sum()) if st_n else 0.0),
            ("Mean absolute deviation (% of capacity)", None, fmt_2dp,
             float(np.abs(st_dev).mean() / cap * 100) if st_n else 0.0),  # formula added after cap_row known
            ("Blocks that incurred a penalty",
             f'=COUNTIF({pen_c}{s}:{pen_c}{e},">0")', fmt_int,
             float((st_pen > 0).sum()) if st_n else 0.0),
            ("Worst single-block penalty (Rs)",
             f"=MAX({pen_c}{s}:{pen_c}{e})", fmt_2dp, float(st_pen.max()) if st_n else 0.0),
            (f"TOTAL STEP {n_} DSM PENALTY FOR THE DAY (Rs)",
             f"=SUM({pen_c}{s}:{pen_c}{e})", fmt_2dp, float(st_pen.sum()) if st_n else 0.0),
        ]
        st_summary_rows_0 = {}
        for i, (label, formula, numfmt, cached) in enumerate(st_defs):
            r0 = st_header_0 + 1 + i
            ws.write(r0, 0, label, label_fmt)
            st_summary_rows_0[label] = r0
            if formula is not None:
                ws.write_formula(r0, 2, formula, numfmt, cached)

        stage_mad_pct_row_0[n_] = st_summary_rows_0["Mean absolute deviation (% of capacity)"]
        st_total_row_0 = st_summary_rows_0[f"TOTAL STEP {n_} DSM PENALTY FOR THE DAY (Rs)"]
        ws.write(st_total_row_0, 0, f"TOTAL STEP {n_} DSM PENALTY FOR THE DAY (Rs)", total_label_fmt)
        ws.write_formula(st_total_row_0, 2, f"=SUM({pen_c}{s}:{pen_c}{e})", total_value_fmt,
                          float(st_pen.sum()) if st_n else 0.0)
        last_summary_row_0 = st_total_row_0

    # ------------------------------------------------------ Slab parameters
    blank2_0 = last_summary_row_0 + 1
    slab_section_0 = blank2_0 + 1
    ws.write(slab_section_0, 0, "DSM SLAB PARAMETERS (the Penalty column references these cells)", section_fmt)

    cap_row_0 = slab_section_0 + 1
    ws.write(cap_row_0, 0, "Installed capacity (MW)", label_fmt)
    ws.write_number(cap_row_0, 2, plant.installed_capacity_mw, input_fmt)
    cap_row = xl_row(cap_row_0)

    energy_factor_row_0 = cap_row_0 + 1
    ws.write(energy_factor_row_0, 0, f"Block energy factor ({plant.block_hours:g} h x 1000 kW/MW)", label_fmt)
    ws.write_number(energy_factor_row_0, 2, plant.block_hours * 1000, input_fmt)
    energy_factor_row = xl_row(energy_factor_row_0)

    slab_table_header_0 = energy_factor_row_0 + 2
    for i, h in enumerate(["Slab", "From %", "To %", "Rate (Rs/kWh)", "Upper edge (MW)"]):
        ws.write(slab_table_header_0, i, h, header_fmt)

    slab_rows: List[int] = []
    for i, slab in enumerate(plant.dsm_slabs):
        r0 = slab_table_header_0 + 1 + i
        r = xl_row(r0)
        slab_rows.append(r)
        ws.write(r0, 0, f"Slab {i + 1}", label_fmt)
        ws.write_number(r0, 1, slab.lower)
        if slab.upper is not None:
            ws.write_number(r0, 2, slab.upper)
        else:
            ws.write_string(r0, 2, "above")
        ws.write_number(r0, 3, slab.rate, input_fmt)
        if slab.upper is not None:
            edge_value = plant.installed_capacity_mw * slab.upper / 100.0
            ws.write_formula(r0, 4, f"=$C${cap_row}*{slab.upper}/100", slab_edge_fmt, edge_value)

    # ---------------------------------------------- Now fill in the formulas
    # that depend on cap_row / energy_factor_row / slab_rows
    for offset in range(n):
        r0 = data_start_0 + offset
        r = xl_row(r0)
        row_i = block_df.iloc[offset]
        is_calc_i = (row_i.get("Status") == "Calculated") if has_status else True

        # F (Deviation %) and G (Penalty) are only live formulas for a
        # Calculated block -- a Pending block gets real blank cells here too,
        # so it can never read as a (wrong) zero deviation/penalty.
        if is_calc_i:
            dev_pct_val = float(row_i["Deviation_%"])
            ws.write_formula(r0, 5, f"=E{r}/$C${cap_row}*100", pct_fmt, dev_pct_val)

            terms = []
            prev_bound = "0"
            for slab, srow in zip(plant.dsm_slabs, slab_rows):
                rate_ref = f"$D${srow}"
                if slab.upper is not None:
                    bound_ref = f"$E${srow}"
                    terms.append(f"MAX(0,MIN(ABS(E{r}),{bound_ref})-{prev_bound})*{rate_ref}")
                    prev_bound = bound_ref
                else:
                    terms.append(f"MAX(0,ABS(E{r})-{prev_bound})*{rate_ref}")
            formula = f"=$C${energy_factor_row}*(" + "+".join(terms) + ")"
            penalty_val = float(row_i["Total_Block_Penalty"])
            ws.write_formula(r0, 6, formula, pen_fmt, penalty_val)
        else:
            ws.write_blank(r0, 5, None, pct_fmt)
            ws.write_blank(r0, 6, None, pen_fmt)

        if has_enercast:
            # Enercast comparison needs BOTH an Enercast forecast AND a real
            # actual meter reading for this block -- if Actual is missing
            # (block Pending), there's nothing to compare Enercast against.
            if pd.notna(row_i.get("Enercast_MW")) and pd.notna(row_i.get("Actual_MW")):
                ws.write_formula(r0, 9, f"=D{r}-I{r}", mw_fmt, float(row_i["Enercast_Deviation_MW"]))
                ws.write_formula(r0, 10, f"=J{r}/$C${cap_row}*100", pct_fmt, float(row_i["Enercast_Deviation_%"]))
                terms_e = []
                prev_bound = "0"
                for slab, srow in zip(plant.dsm_slabs, slab_rows):
                    rate_ref = f"$D${srow}"
                    if slab.upper is not None:
                        bound_ref = f"$E${srow}"
                        terms_e.append(f"MAX(0,MIN(ABS(J{r}),{bound_ref})-{prev_bound})*{rate_ref}")
                        prev_bound = bound_ref
                    else:
                        terms_e.append(f"MAX(0,ABS(J{r})-{prev_bound})*{rate_ref}")
                e_formula = f"=$C${energy_factor_row}*(" + "+".join(terms_e) + ")"
                ws.write_formula(r0, 11, e_formula, pen_fmt, float(row_i["Enercast_Penalty"]))
            else:
                for col in (9, 10, 11):
                    ws.write_blank(r0, col, None, mw_fmt if col == 9 else (pct_fmt if col == 10 else pen_fmt))

        for n_ in active_stages:
            mw_c, dev_c = stage_cols[n_]["mw"], stage_cols[n_]["dev"]
            col0 = stage_col0[n_]
            # This stage's comparison needs BOTH a forecast AND a real
            # actual meter reading for this block -- if Actual is missing
            # (block Pending), there's nothing to compare against.
            if pd.notna(row_i.get(f"Step{n_}_MW")) and pd.notna(row_i.get("Actual_MW")):
                ws.write_formula(r0, col0 + 1, f"=D{r}-{mw_c}{r}", mw_fmt,
                                  float(row_i[f"Step{n_}_Deviation_MW"]))
                ws.write_formula(r0, col0 + 2, f"={dev_c}{r}/$C${cap_row}*100", pct_fmt,
                                  float(row_i[f"Step{n_}_Deviation_%"]))
                terms_st = []
                prev_bound = "0"
                for slab, srow in zip(plant.dsm_slabs, slab_rows):
                    rate_ref = f"$D${srow}"
                    if slab.upper is not None:
                        bound_ref = f"$E${srow}"
                        terms_st.append(f"MAX(0,MIN(ABS({dev_c}{r}),{bound_ref})-{prev_bound})*{rate_ref}")
                        prev_bound = bound_ref
                    else:
                        terms_st.append(f"MAX(0,ABS({dev_c}{r})-{prev_bound})*{rate_ref}")
                st_formula = f"=$C${energy_factor_row}*(" + "+".join(terms_st) + ")"
                ws.write_formula(r0, col0 + 3, st_formula, pen_fmt, float(row_i[f"Step{n_}_Penalty"]))
            else:
                for col in (col0 + 1, col0 + 2, col0 + 3):
                    ws.write_blank(
                        r0, col, None,
                        mw_fmt if col == col0 + 1 else (pct_fmt if col == col0 + 2 else pen_fmt),
                    )

    mad_pct_row_0 = summary_rows_0["Mean absolute deviation (% of capacity)"]
    mad_pct_cached = float(np.abs(dev).mean() / cap * 100) if n_calc else 0.0
    ws.write_formula(
        mad_pct_row_0, 2,
        f"=SUMPRODUCT(ABS(E{s}:E{e}))/COUNT(E{s}:E{e})/$C${cap_row}*100",
        fmt_2dp, mad_pct_cached,
    )

    if has_enercast and e_mad_pct_row_0 is not None:
        e_mask = block_df["Enercast_MW"].notna() & block_df["Actual_MW"].notna()
        e_n = int(e_mask.sum())
        e_dev_for_pct = block_df.loc[e_mask, "Enercast_Deviation_MW"].to_numpy() if e_n else np.array([])
        e_mad_pct_cached = float(np.abs(e_dev_for_pct).mean() / cap * 100) if e_n else 0.0
        ws.write_formula(
            e_mad_pct_row_0, 2,
            f"=SUMPRODUCT(ABS(J{s}:J{e}))/COUNT(J{s}:J{e})/$C${cap_row}*100",
            fmt_2dp, e_mad_pct_cached,
        )

    for n_ in active_stages:
        st_row_0 = stage_mad_pct_row_0.get(n_)
        if st_row_0 is None:
            continue
        dev_c = stage_cols[n_]["dev"]
        st_mask = block_df[f"Step{n_}_MW"].notna() & block_df["Actual_MW"].notna()
        st_n = int(st_mask.sum())
        st_dev_for_pct = block_df.loc[st_mask, f"Step{n_}_Deviation_MW"].to_numpy() if st_n else np.array([])
        st_mad_pct_cached = float(np.abs(st_dev_for_pct).mean() / cap * 100) if st_n else 0.0
        ws.write_formula(
            st_row_0, 2,
            f"=SUMPRODUCT(ABS({dev_c}{s}:{dev_c}{e}))/COUNT({dev_c}{s}:{dev_c}{e})/$C${cap_row}*100",
            fmt_2dp, st_mad_pct_cached,
        )

    # --------------------------------------------------------- Conditional
    if n:
        ws.conditional_format(s - 1, 4, e - 1, 5, {
            "type": "formula",
            "criteria": f"AND(ISNUMBER($E{s}),ABS($E{s})>0.5)",
            "format": red_font,
        })
        ws.conditional_format(s - 1, 4, e - 1, 5, {
            "type": "formula",
            "criteria": f"AND(ISNUMBER($E{s}),ABS($E{s})<=0.5)",
            "format": blue_font,
        })
        ws.conditional_format(s - 1, 6, e - 1, 6, {
            "type": "formula",
            "criteria": f"AND(ISNUMBER($G{s}),$G{s}>0)",
            "format": red_font,
        })
        if has_enercast:
            ws.conditional_format(s - 1, 9, e - 1, 10, {
                "type": "formula",
                "criteria": f"AND(ISNUMBER($J{s}),ABS($J{s})>0.5)",
                "format": red_font,
            })
            ws.conditional_format(s - 1, 9, e - 1, 10, {
                "type": "formula",
                "criteria": f"AND(ISNUMBER($J{s}),ABS($J{s})<=0.5)",
                "format": blue_font,
            })
            ws.conditional_format(s - 1, 11, e - 1, 11, {
                "type": "formula",
                "criteria": f"AND(ISNUMBER($L{s}),$L{s}>0)",
                "format": red_font,
            })
        for n_ in active_stages:
            col0 = stage_col0[n_]
            dev_c, pen_c = stage_cols[n_]["dev"], stage_cols[n_]["pen"]
            ws.conditional_format(s - 1, col0 + 1, e - 1, col0 + 2, {
                "type": "formula",
                "criteria": f"AND(ISNUMBER(${dev_c}{s}),ABS(${dev_c}{s})>0.5)",
                "format": red_font,
            })
            ws.conditional_format(s - 1, col0 + 1, e - 1, col0 + 2, {
                "type": "formula",
                "criteria": f"AND(ISNUMBER(${dev_c}{s}),ABS(${dev_c}{s})<=0.5)",
                "format": blue_font,
            })
            ws.conditional_format(s - 1, col0 + 3, e - 1, col0 + 3, {
                "type": "formula",
                "criteria": f"AND(ISNUMBER(${pen_c}{s}),${pen_c}{s}>0)",
                "format": red_font,
            })

    # -------------------------------------------------------------- Layout
    ws.freeze_panes(data_start_0, 0)
    for i, width in enumerate(col_widths):
        ws.set_column(i, i, width)

    # ------------------------------------------------ Charts (native, on both sheets)
    if n:
        dev_values = block_df["Deviation_MW"].to_numpy()
        pen_values = block_df["Total_Block_Penalty"].to_numpy()
        dev_pct_values = block_df["Deviation_%"].to_numpy()
        e_pen_values = block_df["Enercast_Penalty"].to_numpy() if has_enercast else None
        stage_pen_values = {n_: block_df[f"Step{n_}_Penalty"].to_numpy() for n_ in active_stages}
        stage_pen_col = {n_: stage_cols[n_]["pen"] for n_ in active_stages}
        helper = _write_chart_helper_data(ws_graphs, SHEET, data_start_0, data_end_0, s, e,
                                           dev_values, pen_values, dev_pct_values,
                                           has_enercast, e_pen_values,
                                           active_stages=active_stages,
                                           stage_pen_values=stage_pen_values,
                                           stage_pen_col=stage_pen_col)
        ws_graphs.merge_range(0, 0, 0, 7, f"{plant.name} — DSM Evaluation Charts", title_fmt)
        # Enercast (4 cols) and/or each active stage (4 cols each) and PPA
        # Amount + Status (2 cols) sit to the right of the base columns on
        # the main sheet -- shift the chart anchors so they never sit on top
        # of that data.
        chart_offset = (4 if has_enercast else 0) + 4 * len(active_stages) + 2
        _add_all_charts(wb, ws, SHEET, report_date, data_start_0, data_end_0, helper, dev_values,
                         anchor_col_left=9 + chart_offset, anchor_col_right=23 + chart_offset,
                         row_step=18, size=(430, 260), has_enercast=has_enercast,
                         active_stages=active_stages, stage_col0=stage_col0)
        _add_all_charts(wb, ws_graphs, SHEET, report_date, data_start_0, data_end_0, helper, dev_values,
                         anchor_col_left=0, anchor_col_right=11, row_step=22, size=(560, 340),
                         start_row=2, has_enercast=has_enercast,
                         active_stages=active_stages, stage_col0=stage_col0)

    wb.close()
    return output.getvalue()


# Friendly titles for every native chart, in the same order as the
# dashboard's "Charts" tab, so the main sheet and the Graphs sheet always
# show the same set in the same order as the live dashboard.
CHART_TITLES = {
    "actual_vs_predicted": "Actual vs Predicted Generation",
    "deviation_bar": "Deviation (MW) by Block",
    "penalty_bar": "DSM Penalty (Rs) by Block",
    "cumulative_penalty": "Cumulative DSM Penalty Through the Day",
    "over_under_pie": "Over vs Under Generation Blocks",
    "deviation_histogram": "Distribution of Deviation %",
}

# Helper-data layout on the "Graphs" sheet -- far-right hidden columns that
# feed the charts native Excel can't drive straight from the main sheet's own
# columns (a running cumulative total, category counts, histogram bins).
_H_CUM = 28          # AC: cumulative penalty, one row per data row (aligned with main sheet rows)
_H_PIE_LABEL = 30    # AE: 3 rows -- Over / Under / No Deviation
_H_PIE_COUNT = 31    # AF
_H_ECUM = 29         # AD: Enercast cumulative penalty (comparison-only, when present)
_H_S1CUM = 32        # AG: Step 1 cumulative penalty (comparison-only, when present)
_H_S2CUM = 33        # AH: Step 2 cumulative penalty (comparison-only, when present)
_H_STAGE_CUM = {1: _H_S1CUM, 2: _H_S2CUM}
_H_BIN_EDGE = 34     # AI: 11 histogram bin edges
_H_BIN_LABEL = 35    # AJ: 10 histogram bin labels
_H_BIN_COUNT = 36    # AK: 10 histogram bin counts


def _write_chart_helper_data(ws_graphs, main_sheet: str, data_start_0: int, data_end_0: int, s: int, e: int,
                              dev_values, pen_values, dev_pct_values,
                              has_enercast: bool = False, e_pen_values=None,
                              active_stages=None, stage_pen_values: dict = None,
                              stage_pen_col: dict = None) -> dict:
    """Write hidden formula-driven helper data on the Graphs sheet that the
    native charts reference for series the main sheet's own columns can't
    drive directly: a running cumulative-penalty total, over/under/no-
    deviation counts, and Deviation % histogram bins.

    Every cell is written with a CACHED VALUE (computed in Python from the
    same numbers already in block_df) alongside its live formula -- exactly
    like the main sheet's formula cells -- so the charts that reference these
    cells show correct data immediately in every viewer (LibreOffice, a
    quick-preview pane, pandas) without requiring a formula recalculation
    pass first. Opening the file in a real spreadsheet still shows (and can
    recalculate) the live formula.
    """
    dev_range = f"'{main_sheet}'!$E${s}:$E${e}"
    dev_pct_range = f"'{main_sheet}'!$F${s}:$F${e}"

    # Running cumulative penalty, one formula per data row, aligned to the
    # SAME row numbers as the main sheet so the main sheet's own Time_Label
    # column (B) can be reused as the category axis. nancumsum treats a
    # Pending block's NaN penalty as a zero *contribution* to the running
    # total (matching what SUM() does with a blank G cell) without ever
    # claiming that block itself had a real, calculated penalty of zero.
    cum = np.nancumsum(pen_values) if len(pen_values) else np.array([])
    for offset in range(data_end_0 - data_start_0 + 1):
        r0 = data_start_0 + offset
        r = r0 + 1
        ws_graphs.write_formula(r0, _H_CUM, f"=SUM('{main_sheet}'!$G${s}:$G${r})", None, float(cum[offset]))

    # Enercast running cumulative penalty (comparison only), same pattern as
    # above but summing column L. SUM() naturally skips blank cells for
    # blocks with no Enercast coverage, so partial coverage is handled fine.
    if has_enercast and e_pen_values is not None:
        e_cum = np.nancumsum(np.nan_to_num(e_pen_values, nan=0.0)) if len(e_pen_values) else np.array([])
        for offset in range(data_end_0 - data_start_0 + 1):
            r0 = data_start_0 + offset
            r = r0 + 1
            ws_graphs.write_formula(r0, _H_ECUM, f"=SUM('{main_sheet}'!$L${s}:$L${r})", None, float(e_cum[offset]))

    # Per-stage running cumulative penalty (comparison only), same pattern as
    # Enercast above but summing each active stage's own Penalty column.
    active_stages = active_stages or []
    stage_pen_values = stage_pen_values or {}
    stage_pen_col = stage_pen_col or {}
    stage_cum_ranges = {}
    for n_ in active_stages:
        h_col = _H_STAGE_CUM.get(n_)
        pen_col = stage_pen_col.get(n_)
        pen_vals = stage_pen_values.get(n_)
        if h_col is None or pen_col is None or pen_vals is None:
            continue
        st_cum = np.nancumsum(np.nan_to_num(pen_vals, nan=0.0)) if len(pen_vals) else np.array([])
        for offset in range(data_end_0 - data_start_0 + 1):
            r0 = data_start_0 + offset
            r = r0 + 1
            ws_graphs.write_formula(
                r0, h_col,
                f"=SUM('{main_sheet}'!${pen_col}${s}:${pen_col}${r})",
                None, float(st_cum[offset]),
            )
        stage_cum_ranges[n_] = [None, data_start_0, h_col, data_end_0, h_col]

    # Over / Under / No Deviation counts, for the pie chart.
    pie_defs = [
        ("Over Generation", ">0", int((dev_values > 0).sum())),
        ("Under Generation", "<0", int((dev_values < 0).sum())),
        ("No Deviation", "=0", int((dev_values == 0).sum())),
    ]
    for i, (label, criteria, count) in enumerate(pie_defs):
        ws_graphs.write(i, _H_PIE_LABEL, label)
        ws_graphs.write_formula(i, _H_PIE_COUNT, f'=COUNTIF({dev_range},"{criteria}")', None, count)

    # 10 equal-width Deviation % histogram bins, spanning the data's own
    # min/max, for the histogram chart.
    n_bins = 10
    valid_pct = dev_pct_values[~np.isnan(dev_pct_values)] if len(dev_pct_values) else np.array([])
    if len(valid_pct):
        lo_val, hi_val = float(np.min(valid_pct)), float(np.max(valid_pct))
    else:
        lo_val, hi_val = 0.0, 0.0
    edges = [lo_val + i * (hi_val - lo_val) / n_bins for i in range(n_bins + 1)]
    for i, edge_val in enumerate(edges):
        ws_graphs.write_formula(
            i, _H_BIN_EDGE,
            f"=MIN({dev_pct_range})+({i})*(MAX({dev_pct_range})-MIN({dev_pct_range}))/{n_bins}",
            None, edge_val,
        )
    for i in range(n_bins):
        lo_cell = xl_rowcol_to_cell(i, _H_BIN_EDGE)
        hi_cell = xl_rowcol_to_cell(i + 1, _H_BIN_EDGE)
        upper_op = "<=" if i == n_bins - 1 else "<"
        lo_edge, hi_edge = edges[i], edges[i + 1]
        if i == n_bins - 1:
            bin_count = int(((dev_pct_values >= lo_edge) & (dev_pct_values <= hi_edge)).sum()) if len(dev_pct_values) else 0
        else:
            bin_count = int(((dev_pct_values >= lo_edge) & (dev_pct_values < hi_edge)).sum()) if len(dev_pct_values) else 0
        bin_label = f"{lo_edge:.1f} to {hi_edge:.1f}%"
        ws_graphs.write_formula(i, _H_BIN_LABEL, f'=TEXT({lo_cell},"0.0")&" to "&TEXT({hi_cell},"0.0")&"%"',
                                 None, bin_label)
        ws_graphs.write_formula(
            i, _H_BIN_COUNT,
            f'=COUNTIFS({dev_pct_range},">="&{lo_cell},{dev_pct_range},"{upper_op}"&{hi_cell})',
            None, bin_count,
        )

    ws_graphs.set_column(_H_CUM, _H_BIN_COUNT, None, None, {"hidden": True})

    GRAPHS = "Graphs"
    for n_ in stage_cum_ranges:
        stage_cum_ranges[n_][0] = GRAPHS
    return {
        "cum_values": [GRAPHS, data_start_0, _H_CUM, data_end_0, _H_CUM],
        "ecum_values": [GRAPHS, data_start_0, _H_ECUM, data_end_0, _H_ECUM] if has_enercast else None,
        "stage_cum_values": stage_cum_ranges,
        "pie_categories": [GRAPHS, 0, _H_PIE_LABEL, 2, _H_PIE_LABEL],
        "pie_values": [GRAPHS, 0, _H_PIE_COUNT, 2, _H_PIE_COUNT],
        "hist_categories": [GRAPHS, 0, _H_BIN_LABEL, n_bins - 1, _H_BIN_LABEL],
        "hist_values": [GRAPHS, 0, _H_BIN_COUNT, n_bins - 1, _H_BIN_COUNT],
    }


def _add_all_charts(wb, ws, sheet: str, report_date: str, data_start_0: int, data_end_0: int,
                     helper: dict, dev_values, anchor_col_left: int, anchor_col_right: int, row_step: int,
                     size: tuple, start_row: int = 2, has_enercast: bool = False,
                     active_stages=None, stage_col0: dict = None):
    """Build all 6 native, interactive Excel chart objects and place them on
    `ws` in a two-column grid (left: Actual vs Predicted, Penalty, Over/Under
    pie; right: Deviation, Cumulative penalty, Histogram) -- mirroring the
    dashboard's own Charts tab layout. Every chart is a live Excel chart
    object (hover tooltips, zoom, editable) built from cell references, so
    no image rendering / kaleido is involved at all."""
    cat_range = [sheet, data_start_0, 1, data_end_0, 1]  # column B (Time block)
    width, height = size
    active_stages = active_stages or []
    stage_col0 = stage_col0 or {}
    has_stages = bool(active_stages)

    def _placed(chart, col, row):
        chart.set_size({"width": width, "height": height})
        ws.insert_chart(row, col, chart)

    # 1) Actual vs Predicted -- line chart
    c1 = wb.add_chart({"type": "line"})
    c1.add_series({
        "name": [sheet, 3, 2], "categories": cat_range,
        "values": [sheet, data_start_0, 2, data_end_0, 2],
        "line": {"color": "#15803D", "width": 2},
    })
    c1.add_series({
        "name": [sheet, 3, 3], "categories": cat_range,
        "values": [sheet, data_start_0, 3, data_end_0, 3],
        "line": {"color": "#F59E0B", "width": 2},
    })
    if has_enercast:
        c1.add_series({
            "name": [sheet, 3, 8], "categories": cat_range,
            "values": [sheet, data_start_0, 8, data_end_0, 8],
            "line": {"color": COLOR_ENERCAST, "width": 2, "dash_type": "round_dot"},
        })
    _dash_types = {1: "dash", 2: "dash_dot"}
    for n_ in active_stages:
        col0 = stage_col0.get(n_)
        if col0 is None:
            continue
        c1.add_series({
            "name": [sheet, 3, col0], "categories": cat_range,
            "values": [sheet, data_start_0, col0, data_end_0, col0],
            "line": {"color": STEP_COLOR.get(n_, COLOR_STEP1), "width": 2,
                     "dash_type": _dash_types.get(n_, "dash")},
        })
    c1.set_title({"name": f"Actual vs Predicted Generation - {report_date}"})
    c1.set_x_axis({"name": "Time block", "num_font": {"rotation": -45, "size": 7}})
    c1.set_y_axis({"name": "MW"})
    c1.set_legend({"position": "bottom"})

    # 2) Deviation (MW) by Block -- column chart, red/green per point
    c2 = wb.add_chart({"type": "column"})
    points = [{"fill": {"color": "#16A34A" if v >= 0 else "#DC2626"}} for v in dev_values] or None
    c2.add_series({
        "name": "Our Deviation (MW)", "categories": cat_range,
        "values": [sheet, data_start_0, 4, data_end_0, 4],
        "points": points,
    })
    if has_enercast:
        c2.add_series({
            "name": [sheet, 3, 9], "categories": cat_range,
            "values": [sheet, data_start_0, 9, data_end_0, 9],
            "fill": {"color": COLOR_ENERCAST},
        })
    for n_ in active_stages:
        col0 = stage_col0.get(n_)
        if col0 is None:
            continue
        c2.add_series({
            "name": [sheet, 3, col0 + 1], "categories": cat_range,
            "values": [sheet, data_start_0, col0 + 1, data_end_0, col0 + 1],
            "fill": {"color": STEP_COLOR.get(n_, COLOR_STEP1)},
        })
    c2.set_title({"name": f"Deviation (MW) by Block - {report_date}"})
    c2.set_x_axis({"name": "Time block", "num_font": {"rotation": -45, "size": 7}})
    c2.set_y_axis({"name": "Deviation MW"})
    c2.set_legend({"none": not (has_enercast or has_stages)})

    # 3) DSM Penalty (Rs) by Block -- column chart
    c3 = wb.add_chart({"type": "column"})
    c3.add_series({
        "name": "Our Penalty (Rs)", "categories": cat_range,
        "values": [sheet, data_start_0, 6, data_end_0, 6],
        "fill": {"color": "#7C3AED"},
    })
    if has_enercast:
        c3.add_series({
            "name": [sheet, 3, 11], "categories": cat_range,
            "values": [sheet, data_start_0, 11, data_end_0, 11],
            "fill": {"color": COLOR_ENERCAST},
        })
    for n_ in active_stages:
        col0 = stage_col0.get(n_)
        if col0 is None:
            continue
        c3.add_series({
            "name": [sheet, 3, col0 + 3], "categories": cat_range,
            "values": [sheet, data_start_0, col0 + 3, data_end_0, col0 + 3],
            "fill": {"color": STEP_COLOR.get(n_, COLOR_STEP1)},
        })
    c3.set_title({"name": f"DSM Penalty (Rs) by Block - {report_date}"})
    c3.set_x_axis({"name": "Time block", "num_font": {"rotation": -45, "size": 7}})
    c3.set_y_axis({"name": "Penalty (Rs)"})
    c3.set_legend({"none": not (has_enercast or has_stages)})

    # 4) Cumulative DSM Penalty -- area chart
    c4 = wb.add_chart({"type": "area"})
    c4.add_series({
        "name": "Our Cumulative Penalty", "categories": cat_range,
        "values": helper["cum_values"],
        "fill": {"color": "#0F766E", "transparency": 30},
        "line": {"color": "#0F766E", "width": 2},
    })
    if has_enercast and helper.get("ecum_values"):
        c4.add_series({
            "name": "Enercast Cumulative Penalty", "categories": cat_range,
            "values": helper["ecum_values"],
            "fill": {"color": COLOR_ENERCAST, "transparency": 40},
            "line": {"color": COLOR_ENERCAST, "width": 2, "dash_type": "round_dot"},
        })
    _stage_cum_values = helper.get("stage_cum_values") or {}
    for n_ in active_stages:
        cum_range = _stage_cum_values.get(n_)
        if not cum_range:
            continue
        color = STEP_COLOR.get(n_, COLOR_STEP1)
        c4.add_series({
            "name": f"Step {n_} Cumulative Penalty", "categories": cat_range,
            "values": cum_range,
            "fill": {"color": color, "transparency": 40},
            "line": {"color": color, "width": 2, "dash_type": _dash_types.get(n_, "dash")},
        })
    c4.set_title({"name": f"Cumulative DSM Penalty Through the Day - {report_date}"})
    c4.set_x_axis({"name": "Time block", "num_font": {"rotation": -45, "size": 7}})
    c4.set_y_axis({"name": "Cumulative Penalty (Rs)"})
    c4.set_legend({"none": not (has_enercast or has_stages)})

    # 5) Over vs Under Generation -- pie chart
    c5 = wb.add_chart({"type": "pie"})
    c5.add_series({
        "categories": helper["pie_categories"], "values": helper["pie_values"],
        "points": [{"fill": {"color": "#16A34A"}}, {"fill": {"color": "#DC2626"}}, {"fill": {"color": "#64748B"}}],
        "data_labels": {"percentage": True, "category": True},
    })
    c5.set_title({"name": "Over vs Under Generation Blocks"})

    # 6) Distribution of Deviation % -- histogram-style column chart
    c6 = wb.add_chart({"type": "column"})
    c6.add_series({
        "name": CHART_TITLES["deviation_histogram"], "categories": helper["hist_categories"],
        "values": helper["hist_values"], "fill": {"color": "#15803D"}, "gap": 0,
    })
    c6.set_title({"name": "Distribution of Deviation %"})
    c6.set_x_axis({"name": "Deviation % of capacity", "num_font": {"size": 7}})
    c6.set_y_axis({"name": "Number of blocks"})
    c6.set_legend({"none": True})

    left = [c1, c3, c5]
    right = [c2, c4, c6]
    for col, charts in ((anchor_col_left, left), (anchor_col_right, right)):
        for i, chart in enumerate(charts):
            _placed(chart, col, start_row + i * row_step)
