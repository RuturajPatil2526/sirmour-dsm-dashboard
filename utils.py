"""
utils.py
========
File loading, validation, column auto-detection and dataset merging for the
AI Schedule Evaluation Dashboard.

Real-world exports rarely match a single fixed CSV schema (comment/metadata
rows at the top, different header spellings, values given in kW instead of
MW, timestamps instead of an explicit block number, ...). Everything in this
module is written to be forgiving of that, while still failing loudly and
clearly when a file genuinely cannot be understood.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

import numpy as np
import pandas as pd

from config import COLUMN_ALIASES, block_time_label


class FileValidationError(Exception):
    """Raised when an uploaded file cannot be safely parsed."""
    pass


@dataclass
class ParseResult:
    df: pd.DataFrame
    warnings: List[str] = field(default_factory=list)
    detected_columns: dict = field(default_factory=dict)
    has_date: bool = True


# ---------------------------------------------------------------------------
# Raw file reading
# ---------------------------------------------------------------------------
def read_uploaded_table(uploaded_file) -> pd.DataFrame:
    """Read a Streamlit UploadedFile / file path / file-like object into a
    raw DataFrame, tolerating leading comment (#) lines and blank lines that
    are common in exported SCADA / forecasting reports.
    """
    name = getattr(uploaded_file, "name", str(uploaded_file))
    is_excel = name.lower().endswith((".xlsx", ".xls"))

    if is_excel:
        try:
            return pd.read_excel(uploaded_file, engine="openpyxl")
        except Exception as exc:  # pragma: no cover
            raise FileValidationError(f"Could not read Excel file '{name}': {exc}")

    # CSV / text path -----------------------------------------------------
    try:
        if hasattr(uploaded_file, "read"):
            raw_bytes = uploaded_file.read()
            if hasattr(uploaded_file, "seek"):
                uploaded_file.seek(0)
            if isinstance(raw_bytes, bytes):
                text = raw_bytes.decode("utf-8-sig", errors="replace")
            else:
                text = raw_bytes
        else:
            with open(uploaded_file, "r", encoding="utf-8-sig", errors="replace") as fh:
                text = fh.read()
    except Exception as exc:
        raise FileValidationError(f"Could not read file '{name}': {exc}")

    if not text.strip():
        raise FileValidationError(f"File '{name}' is empty.")

    try:
        df = pd.read_csv(
            io.StringIO(text),
            comment="#",
            skip_blank_lines=True,
            engine="python",
            sep=None,  # sniff delimiter (handles comma / semicolon / tab)
        )
    except Exception as exc:
        raise FileValidationError(
            f"Could not parse '{name}' as CSV. Please check the file format. Details: {exc}"
        )

    if df.empty or len(df.columns) < 2:
        raise FileValidationError(
            f"'{name}' does not look like a valid data table (no usable rows/columns found)."
        )

    return df


# ---------------------------------------------------------------------------
# Column auto-detection
# ---------------------------------------------------------------------------
def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower())


def find_column(df: pd.DataFrame, category: str) -> Optional[str]:
    aliases = [_norm(a) for a in COLUMN_ALIASES.get(category, [])]
    norm_map = {_norm(c): c for c in df.columns}
    # exact match first
    for alias in aliases:
        if alias in norm_map:
            return norm_map[alias]
    # substring match as a fallback (e.g. "Predicted (MW) @15min")
    for norm_col, orig_col in norm_map.items():
        for alias in aliases:
            if alias and alias in norm_col:
                return orig_col
    return None


_TIME_RANGE_SUFFIX = re.compile(r"\s*-\s*\d{1,2}:\d{2}(:\d{2})?\s*$")


def _clean_timestamp_series(raw: pd.Series) -> pd.Series:
    """Strip a trailing '- HH:MM' (or '- HH:MM:SS') range-end suffix from a
    timestamp-like column before parsing, e.g. '2026-09-11 05:15 - 05:30'
    (a "Time Interval" column marking a block's start AND end) becomes
    '2026-09-11 05:15', which pandas can parse. A plain single timestamp
    with no such suffix passes through unchanged. Only touches string/object
    columns -- a column that's already datetime-typed is returned as-is.
    """
    if not pd.api.types.is_object_dtype(raw) and not pd.api.types.is_string_dtype(raw):
        return raw
    return raw.astype(str).str.replace(_TIME_RANGE_SUFFIX, "", regex=True)


def _extract_block_number(value) -> Optional[int]:
    if pd.isna(value):
        return None
    if isinstance(value, (int, np.integer)):
        return int(value)
    if isinstance(value, float):
        return int(value)
    match = re.search(r"\d+", str(value))
    return int(match.group()) if match else None


# ---------------------------------------------------------------------------
# Normalising a single uploaded file (schedule OR meter) into:
#   Date (date or None) | Block (int) | Time_Label (str) | Value_MW (float)
# ---------------------------------------------------------------------------
def parse_energy_file(
    uploaded_file,
    role: str,
    block_minutes: int = 15,
    timestamp_marks: str = "start",
) -> ParseResult:
    """
    role: 'schedule' or 'meter'
    timestamp_marks: 'start' or 'end' -- whether a timestamp column marks the
        start or the end of its 15-minute delivery block.
    """
    warnings: List[str] = []
    df_raw = read_uploaded_table(uploaded_file)
    df_raw.columns = [str(c).strip() for c in df_raw.columns]

    value_category = "mw_predicted" if role == "schedule" else "mw_actual"
    value_kw_category = "kw_predicted" if role == "schedule" else "kw_actual"
    label = "AI Schedule" if role == "schedule" else "Meter Data"

    col_date = find_column(df_raw, "date")
    col_block = find_column(df_raw, "block")
    col_ts = find_column(df_raw, "timestamp")
    col_val_mw = find_column(df_raw, value_category)
    col_val_kw = find_column(df_raw, value_kw_category)
    # Multi-step AI Schedule format (only relevant for the schedule file
    # itself): "Step 1 Meter Base Forecast MW" -> "Step 2 Weather Adjustment
    # MW" -> "Step 3 Plant Profile Adjustment MW". Whichever of these is the
    # LAST (highest-numbered) stage present in the file becomes the official
    # Scheduled_MW that drives the main DSM report; every earlier stage that
    # is also present is carried alongside as its own, independently
    # penalised comparison column (see calculator.evaluate_step1/
    # evaluate_step2) -- exactly like Enercast, but sourced from the same
    # file. A file with only one step column (or the old single-column
    # format) works exactly as before.
    step_cols = {}
    if role == "schedule":
        for n in (1, 2, 3):
            c = find_column(df_raw, f"mw_step{n}")
            if c is not None:
                step_cols[n] = c

    detected = {
        "date": col_date, "block": col_block, "timestamp": col_ts,
        "value_mw": col_val_mw, "value_kw": col_val_kw,
        "step1_mw": step_cols.get(1), "step2_mw": step_cols.get(2),
        "step3_mw": step_cols.get(3),
    }

    # -- value column -------------------------------------------------
    value_series = None
    comparison_step_series = {}  # {stage_number: pd.Series} for every non-official stage
    if step_cols:
        official_n = max(step_cols)
        official_col = step_cols[official_n]
        value_series = pd.to_numeric(df_raw[official_col], errors="coerce")
        detected["value_mw"] = official_col
        other_ns = sorted(n for n in step_cols if n != official_n)
        if other_ns:
            for n in other_ns:
                comparison_step_series[n] = pd.to_numeric(df_raw[step_cols[n]], errors="coerce")
            stage_word = "step" if len(step_cols) == 1 else f"{len(step_cols)}-step"
            others_desc = ", ".join(f"'{step_cols[n]}' (Step {n})" for n in other_ns)
            warnings.append(
                f"{label}: detected {stage_word} AI Schedule format — using "
                f"'{official_col}' (Step {official_n}, the final stage) as "
                f"the official schedule, and {others_desc} as separate "
                f"comparison-only schedule(s). Each is penalised "
                f"independently against actual meter data."
            )
    elif col_val_mw is not None:
        value_series = pd.to_numeric(df_raw[col_val_mw], errors="coerce")
    elif col_val_kw is not None:
        value_series = pd.to_numeric(df_raw[col_val_kw], errors="coerce") / 1000.0
        warnings.append(
            f"{label}: no MW column found — converted '{col_val_kw}' from kW to MW."
        )
    else:
        # last-resort fallback: any single numeric column with 'mw' or 'power' in the name
        candidates = [
            c for c in df_raw.columns
            if re.search(r"mw|power|kw", _norm(c)) and pd.api.types.is_numeric_dtype(
                pd.to_numeric(df_raw[c], errors="coerce")
            )
        ]
        if len(candidates) == 1:
            value_series = pd.to_numeric(df_raw[candidates[0]], errors="coerce")
            if "kw" in _norm(candidates[0]):
                value_series = value_series / 1000.0
            warnings.append(
                f"{label}: could not confidently identify the value column — "
                f"used '{candidates[0]}' as the generation value."
            )
        else:
            raise FileValidationError(
                f"Could not find a Predicted/Actual MW (or kW) column in the {label} file. "
                f"Found columns: {list(df_raw.columns)}"
            )

    # -- date / block -------------------------------------------------
    has_date = True
    if col_block is not None:
        block_series = df_raw[col_block].map(_extract_block_number)
        if col_date is not None:
            date_series = pd.to_datetime(df_raw[col_date], errors="coerce").dt.date
        elif col_ts is not None:
            date_series = pd.to_datetime(_clean_timestamp_series(df_raw[col_ts]), errors="coerce").dt.date
            if date_series.isna().all():
                # The timestamp column couldn't be parsed at all (e.g. an
                # unrecognised format) -- fall back to Block-only matching
                # rather than silently merging on a column that's entirely
                # NaT, which would otherwise raise a confusing dtype error
                # downstream in merge_datasets().
                date_series = pd.Series([None] * len(df_raw))
                has_date = False
                warnings.append(
                    f"{label}: could not parse dates from column '{col_ts}' — "
                    f"matching will be done on Block number only."
                )
        else:
            date_series = pd.Series([None] * len(df_raw))
            has_date = False
            warnings.append(
                f"{label}: no date column found — matching will be done on Block number only."
            )
    elif col_ts is not None:
        ts_series = pd.to_datetime(_clean_timestamp_series(df_raw[col_ts]), errors="coerce")
        if ts_series.isna().all():
            raise FileValidationError(
                f"Could not parse any timestamps in column '{col_ts}' of the {label} file."
            )
        minutes = ts_series.dt.hour * 60 + ts_series.dt.minute
        if timestamp_marks == "end":
            # a timestamp of 00:00 marking block-end belongs to block 96 of
            # the *previous* minute bucket -> shift back by one block width
            minutes = (minutes - block_minutes) % (24 * 60)
        block_series = (minutes // block_minutes + 1).astype("Int64")
        date_series = ts_series.dt.date
        if col_date is not None:
            # explicit date column takes precedence if present
            date_series = pd.to_datetime(df_raw[col_date], errors="coerce").dt.date
    elif col_date is not None:
        raise FileValidationError(
            f"The {label} file has a Date column but no Block number or Time column, "
            f"so 15-minute blocks cannot be determined."
        )
    else:
        raise FileValidationError(
            f"The {label} file needs either a 'Block' column or a 'Time/TimeStamp' "
            f"column so blocks can be identified. Found columns: {list(df_raw.columns)}"
        )

    data = {
        "Date": date_series,
        "Block": block_series,
        "Value_MW": value_series,
    }
    if role == "schedule":
        col_gen_at = find_column(df_raw, "generated_at")
        detected["generated_at"] = col_gen_at
        if col_gen_at is not None:
            data["Generated_At"] = df_raw[col_gen_at].astype(str)
        for n, series in comparison_step_series.items():
            data[f"Step{n}_MW"] = series

    out = pd.DataFrame(data)

    out = out.dropna(subset=["Block", "Value_MW"]).copy()
    out["Block"] = out["Block"].astype(int)
    out["Time_Label"] = out["Block"].apply(lambda b: block_time_label(b, block_minutes))

    n_before = len(out)
    dup_mask = out.duplicated(subset=["Date", "Block"], keep="first")
    n_dups = int(dup_mask.sum())
    if n_dups:
        warnings.append(
            f"{label}: found {n_dups} duplicate block entr{'y' if n_dups == 1 else 'ies'} "
            f"— kept the first occurrence of each and discarded the rest."
        )
        out = out[~dup_mask]

    if out.empty:
        raise FileValidationError(f"No valid rows remained in the {label} file after cleaning.")

    rename = "Scheduled_MW" if role == "schedule" else "Actual_MW"
    out = out.rename(columns={"Value_MW": rename})

    return ParseResult(df=out, warnings=warnings, detected_columns=detected, has_date=has_date)


# ---------------------------------------------------------------------------
# Merge schedule + meter data
# ---------------------------------------------------------------------------
def merge_datasets(
    schedule_result: ParseResult, meter_result: ParseResult,
    block_minutes: int = 15, blocks_per_day: int = 96,
) -> Tuple[pd.DataFrame, List[str]]:
    """Match the AI Schedule and Meter Data files block-by-block.

    IMPORTANT: every block from 1..blocks_per_day is always present in the
    returned dataframe (for every date found in the files) -- a block whose
    schedule and/or meter value is missing is NEVER dropped and NEVER
    silently treated as zero. It is kept as a row with NaN in
    Scheduled_MW/Actual_MW; calculator.evaluate_schedule() marks such rows
    Status="Pending" (penalty = null) rather than computing a (wrong) number
    for them.
    """
    warnings = list(schedule_result.warnings) + list(meter_result.warnings)

    sched = schedule_result.df.copy()
    meter = meter_result.df.copy()

    use_date = schedule_result.has_date and meter_result.has_date
    if not use_date:
        warnings.append(
            "Matching performed on Block number only (date missing from one or both files)."
        )
        sched = sched.drop(columns=["Date"], errors="ignore")
        meter = meter.drop(columns=["Date"], errors="ignore")

    # Build the full day's block skeleton (1..blocks_per_day) for every date
    # seen in either file, so every possible block is represented -- even
    # ones neither file has any data for at all.
    if use_date:
        dates = sorted(set(sched["Date"].dropna().unique()) | set(meter["Date"].dropna().unique()))
        if not dates:
            dates = [None]
    else:
        dates = [None]

    blocks_per_day = max(1, int(blocks_per_day))
    skeleton = pd.DataFrame(
        [
            {"Date": d, "Block": b, "Time_Label": block_time_label(b, block_minutes)}
            for d in dates for b in range(1, blocks_per_day + 1)
        ]
    )
    if not use_date:
        skeleton = skeleton.drop(columns=["Date"])

    keys = ["Date", "Block"] if use_date else ["Block"]

    # Drop each source's own Time_Label -- the skeleton's is authoritative
    # for EVERY block (including ones neither file covers).
    sched_slim = sched.drop(columns=["Time_Label"], errors="ignore")
    meter_slim = meter.drop(columns=["Time_Label"], errors="ignore")

    merged = skeleton.merge(sched_slim, on=keys, how="left")
    merged = merged.merge(meter_slim, on=keys, how="left", suffixes=("", "_meter"))

    if "Date" not in merged.columns:
        merged["Date"] = None

    merged = merged.sort_values(["Block"] if not use_date else ["Date", "Block"]).reset_index(drop=True)

    # final sanity: numeric coercion (NaN stays NaN -- never filled with 0)
    merged["Scheduled_MW"] = pd.to_numeric(merged["Scheduled_MW"], errors="coerce")
    merged["Actual_MW"] = pd.to_numeric(merged["Actual_MW"], errors="coerce")
    for n in (1, 2, 3):
        col = f"Step{n}_MW"
        if col in merged.columns:
            merged[col] = pd.to_numeric(merged[col], errors="coerce")

    missing_sched = merged["Scheduled_MW"].isna()
    missing_meter = merged["Actual_MW"].isna()
    n_pending = int((missing_sched | missing_meter).sum())
    n_calculated = int((~missing_sched & ~missing_meter).sum())

    if n_calculated == 0:
        raise FileValidationError(
            "No matching blocks were found between the AI Schedule and Meter Data files. "
            "Please check that both files cover the same date and use the same block numbering."
        )

    if n_pending:
        warnings.append(
            f"{n_pending} of {len(merged)} block(s) had missing AI Schedule and/or Meter "
            f"data — kept as 'Pending' (penalty = null, never treated as zero)."
        )

    cols = ["Date", "Block", "Time_Label", "Scheduled_MW", "Actual_MW"]
    for n in (1, 2, 3):
        col = f"Step{n}_MW"
        if col in merged.columns:
            cols.append(col)
    if "Generated_At" in merged.columns:
        cols.append("Generated_At")
    merged = merged[cols]
    return merged, warnings


# ---------------------------------------------------------------------------
# Optional third file: Enercast (or any other third-party forecast), shown
# alongside our own AI Schedule purely for comparison -- never used to grade
# our own forecast, which is always graded against actual meter data only.
# ---------------------------------------------------------------------------
def parse_enercast_file(
    uploaded_file, block_minutes: int = 15, timestamp_marks: str = "start"
) -> ParseResult:
    """Parse a third-party forecast file (e.g. Enercast's own predicted-MW
    export). Reuses the same column auto-detection as the AI Schedule file
    (role='schedule' picks up 'Scheduled MW' / 'Predicted (MW)' etc. aliases)
    but is matched to the main dataset by BLOCK NUMBER ONLY -- these exports
    commonly carry no reliable date column of their own (just a Time-of-day
    column), so trusting their date would silently break the match.
    """
    result = parse_energy_file(uploaded_file, role="schedule",
                                block_minutes=block_minutes, timestamp_marks=timestamp_marks)
    df = result.df.rename(columns={"Scheduled_MW": "Enercast_MW"})
    df = df[["Block", "Enercast_MW"]].drop_duplicates(subset=["Block"], keep="first")
    result.df = df
    return result


def merge_enercast(block_df: pd.DataFrame, enercast_result: ParseResult) -> Tuple[pd.DataFrame, List[str]]:
    """Left-merge the (already schedule+meter matched) block_df with the
    optional Enercast data, matching on Block number only. Blocks the
    Enercast file doesn't cover simply get a blank Enercast_MW -- this never
    excludes a block from the main DSM calculation, since Enercast is
    comparison-only."""
    warnings = list(enercast_result.warnings)
    merged = block_df.merge(enercast_result.df, on="Block", how="left")
    covered = int(merged["Enercast_MW"].notna().sum())
    missing = len(merged) - covered
    if missing:
        warnings.append(
            f"Enercast: {missing} of {len(merged)} block(s) had no matching Enercast data "
            f"(left blank in the 'Enercast (MW)' column)."
        )
    return merged, warnings
