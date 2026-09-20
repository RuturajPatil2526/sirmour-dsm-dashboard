"""
calculator.py
=============
Core DSM (Deviation Settlement Mechanism) penalty engine.

All formulas below follow the plant's business rules exactly:

    Deviation MW        = Actual MW - Scheduled MW
    Deviation %         = (Deviation MW / Installed Capacity) x 100
    Block Energy (kWh)  = Deviation MW x Block Hours x 1000        (block hours = 0.25 for 15 min)

    Penalty is charged using PIECEWISE SLAB logic on the *absolute* deviation
    percentage: only the portion of the deviation that falls inside a given
    slab is charged at that slab's rate.

    For a slab [lo, hi):
        slab_width_covered = overlap between [lo, hi] and [0, abs(Deviation %)]
        Energy in slab      = |Block Energy| x slab_width_covered / abs(Deviation %)
        Penalty for slab    = Energy in slab x slab rate

    Total Block Penalty = sum of penalties across all slabs.

This module is deliberately agnostic of Streamlit / plotting / Excel -- it
only depends on pandas/numpy and config.PlantConfig, so it can be unit
tested and reused independently.
"""

from __future__ import annotations

from typing import List, Optional

import numpy as np
import pandas as pd

from config import PlantConfig, DSMSlab


def _slab_columns(slabs: List[DSMSlab]) -> List[str]:
    return [f"Energy_{s.label}" for s in slabs], [f"Penalty_{s.label}" for s in slabs]


def compute_block_penalty(abs_dev_pct: float, block_energy_abs_kwh: float, slabs: List[DSMSlab]):
    """Return (energy_per_slab: list[float], penalty_per_slab: list[float], total_penalty: float)
    for a single block, using the piecewise slab allocation described above.
    """
    n = len(slabs)
    energies = [0.0] * n
    penalties = [0.0] * n

    if abs_dev_pct <= 0 or block_energy_abs_kwh <= 0:
        return energies, penalties, 0.0

    for i, slab in enumerate(slabs):
        lo = slab.lower
        hi = slab.upper if slab.upper is not None else abs_dev_pct
        overlap = max(0.0, min(hi, abs_dev_pct) - lo)
        if overlap <= 0:
            continue
        slab_energy = block_energy_abs_kwh * (overlap / abs_dev_pct)
        slab_penalty = slab_energy * slab.rate
        energies[i] = slab_energy
        penalties[i] = slab_penalty

    total_penalty = float(np.sum(penalties))
    return energies, penalties, total_penalty


def evaluate_schedule(merged_df: pd.DataFrame, plant: PlantConfig) -> pd.DataFrame:
    """Take the merged (Date, Block, Time_Label, Scheduled_MW, Actual_MW)
    dataframe -- which now always spans every block of the day (1..N), with
    NaN Scheduled_MW/Actual_MW for any block that had no matching data -- and
    compute every required DSM column, block by block.

    A block with a missing schedule and/or meter value is NEVER converted to
    zero and NEVER silently skipped: it gets Status="Pending" and every
    computed column (Deviation, Penalty, ...) is left as NaN ("null") for
    that block. Only blocks with BOTH a schedule and a meter value get
    Status="Calculated" and a real (possibly zero) penalty. Under- and
    over-generation are always charged identically -- this uses the standard
    absolute-deviation DSM slab logic, never a separate payable/receivable
    settlement split.
    """
    df = merged_df.copy().reset_index(drop=True)
    capacity = plant.installed_capacity_mw
    block_hours = plant.block_hours
    slabs = plant.dsm_slabs
    ppa_rate = getattr(plant, "ppa_rate", 0.0)

    has_sched = df["Scheduled_MW"].notna()
    has_actual = df["Actual_MW"].notna()
    is_calculated = has_sched & has_actual

    df["Deviation_MW"] = np.where(is_calculated, df["Actual_MW"] - df["Scheduled_MW"], np.nan)
    df["Deviation_%"] = np.where(is_calculated, df["Deviation_MW"] / capacity * 100.0, np.nan)
    df["Abs_Deviation_%"] = np.abs(df["Deviation_%"])
    df["Block_Energy_kWh"] = np.where(is_calculated, df["Deviation_MW"] * block_hours * 1000.0, np.nan)

    energy_cols, penalty_cols = _slab_columns(slabs)
    for c in energy_cols + penalty_cols:
        df[c] = np.nan

    total_penalty = np.full(len(df), np.nan)

    for idx in df.index[is_calculated]:
        abs_pct = df.at[idx, "Abs_Deviation_%"]
        energy_abs = abs(df.at[idx, "Block_Energy_kWh"])
        energies, penalties, tot = compute_block_penalty(abs_pct, energy_abs, slabs)
        for c, v in zip(energy_cols, energies):
            df.at[idx, c] = v
        for c, v in zip(penalty_cols, penalties):
            df.at[idx, c] = v
        total_penalty[idx] = tot

    df["Total_Block_Penalty"] = total_penalty
    df["Status"] = np.where(is_calculated, "Calculated", "Pending")
    df["Deviation_Type"] = np.select(
        [~is_calculated, df["Deviation_MW"] > 1e-9, df["Deviation_MW"] < -1e-9],
        ["Pending", "Over Generation", "Under Generation"],
        default="No Deviation",
    )

    # PPA Amount (Rs) -- display-only reference figure, computed from the
    # scheduled energy whenever a schedule value exists, independent of
    # whether meter data has arrived yet for that block. Never affects the
    # DSM penalty above.
    scheduled_kwh = df["Scheduled_MW"] * block_hours * 1000.0
    df["PPA_Amount"] = np.where(has_sched, scheduled_kwh * ppa_rate, np.nan)

    ordered_cols = (
        ["Date", "Block", "Time_Label", "Scheduled_MW", "Actual_MW",
         "Deviation_MW", "Deviation_%", "Abs_Deviation_%", "Block_Energy_kWh"]
        + energy_cols + penalty_cols
        + ["Total_Block_Penalty", "Deviation_Type", "Status", "PPA_Amount"]
    )
    if "Generated_At" in df.columns:
        ordered_cols.append("Generated_At")
    # Step1_MW / Step2_MW (present only when the AI Schedule file used the
    # multi-step format and that stage isn't the official one) are carried
    # through here too, so calculator.evaluate_step1()/evaluate_step2() and
    # everything downstream (charts, Excel export) can still see them --
    # they're otherwise not touched by this function at all.
    for n in (1, 2):
        col = f"Step{n}_MW"
        if col in df.columns:
            ordered_cols.append(col)
    df = df[ordered_cols]
    return df


def evaluate_enercast(df: pd.DataFrame, plant: PlantConfig) -> pd.DataFrame:
    """Add Enercast comparison columns (Enercast_Deviation_MW,
    Enercast_Deviation_%, Enercast_Penalty) using the SAME piecewise DSM
    slab logic as evaluate_schedule(), but graded against Enercast_MW
    instead of our own Scheduled_MW.

    IMPORTANT: this is comparison-only. Our own forecast (Scheduled_MW) is
    always graded against actual meter data alone -- Enercast never affects
    Total_Block_Penalty or any other column evaluate_schedule() already
    computed. Blocks with no Enercast data (Enercast_MW is NaN) simply get
    NaN in every Enercast_* column here, and are excluded from any Enercast
    summary totals.

    A no-op (returns df unchanged) if the "Enercast_MW" column isn't present
    -- callers only invoke this when an Enercast file was actually uploaded.
    """
    if "Enercast_MW" not in df.columns:
        return df

    df = df.copy()
    capacity = plant.installed_capacity_mw
    block_hours = plant.block_hours
    slabs = plant.dsm_slabs

    # Enercast is graded against the actual meter reading too -- if the meter
    # value itself is missing (block Pending), there is nothing to compare
    # Enercast against, so leave it NaN rather than treating the missing
    # actual as zero.
    has_enercast = df["Enercast_MW"].notna() & df["Actual_MW"].notna()
    dev_mw = (df["Actual_MW"] - df["Enercast_MW"]).where(has_enercast)
    dev_pct = (dev_mw / capacity * 100.0)
    block_energy = (dev_mw * block_hours * 1000.0)

    penalties = np.full(len(df), np.nan)
    for idx in df.index[has_enercast]:
        abs_pct = abs(dev_pct.loc[idx])
        energy_abs = abs(block_energy.loc[idx])
        _, _, tot = compute_block_penalty(abs_pct, energy_abs, slabs)
        penalties[df.index.get_loc(idx)] = tot

    df["Enercast_Deviation_MW"] = dev_mw
    df["Enercast_Deviation_%"] = dev_pct
    df["Enercast_Penalty"] = penalties
    return df


def enercast_summary(df: pd.DataFrame) -> Optional[dict]:
    """Head-to-head roll-up (blocks both forecasts covered) -- our own
    Total_Block_Penalty vs Enercast_Penalty, and mean absolute deviation for
    each. Returns None if no Enercast data is present."""
    if "Enercast_MW" not in df.columns:
        return None
    covered = df[df["Enercast_Deviation_MW"].notna()] if "Enercast_Deviation_MW" in df.columns \
        else df[df["Enercast_MW"].notna() & df["Actual_MW"].notna()]
    n = len(covered)
    if n == 0:
        return None
    return {
        "Blocks compared": n,
        "Our Total Penalty (Rs, same blocks)": float(covered["Total_Block_Penalty"].sum()),
        "Enercast Total Penalty (Rs)": float(covered["Enercast_Penalty"].sum()),
        "Our Mean Abs Deviation (MW)": float(covered["Deviation_MW"].abs().mean()),
        "Enercast Mean Abs Deviation (MW)": float(covered["Enercast_Deviation_MW"].abs().mean()),
    }


def _evaluate_comparison_stage(df: pd.DataFrame, plant: PlantConfig, mw_col: str, prefix: str) -> pd.DataFrame:
    """Shared implementation behind evaluate_step1() / evaluate_step2():
    add `{prefix}_Deviation_MW`, `{prefix}_Deviation_%` and `{prefix}_Penalty`
    columns, using the SAME piecewise DSM slab logic as evaluate_schedule(),
    but graded against `mw_col` instead of the official Scheduled_MW.

    This lets an earlier stage of a multi-step AI forecasting pipeline (e.g.
    Step 1's raw base forecast, or Step 2's weather-adjusted forecast, when
    Step 3 is the official/final schedule) be penalised independently, so
    the value each later stage adds (or costs) can be seen directly. It is
    comparison-only -- it never affects Total_Block_Penalty or any other
    column evaluate_schedule() already computed, exactly the same
    relationship Enercast has with the official schedule.

    A no-op (returns df unchanged) if `mw_col` isn't present -- callers
    only invoke this when the AI Schedule file actually had that stage.
    """
    if mw_col not in df.columns:
        return df

    df = df.copy()
    capacity = plant.installed_capacity_mw
    block_hours = plant.block_hours
    slabs = plant.dsm_slabs

    # Graded against the actual meter reading too -- if Actual is missing
    # (block Pending), there is nothing to compare this stage against.
    has_stage = df[mw_col].notna() & df["Actual_MW"].notna()
    dev_mw = (df["Actual_MW"] - df[mw_col]).where(has_stage)
    dev_pct = (dev_mw / capacity * 100.0)
    block_energy = (dev_mw * block_hours * 1000.0)

    penalties = np.full(len(df), np.nan)
    for idx in df.index[has_stage]:
        abs_pct = abs(dev_pct.loc[idx])
        energy_abs = abs(block_energy.loc[idx])
        _, _, tot = compute_block_penalty(abs_pct, energy_abs, slabs)
        penalties[df.index.get_loc(idx)] = tot

    df[f"{prefix}_Deviation_MW"] = dev_mw
    df[f"{prefix}_Deviation_%"] = dev_pct
    df[f"{prefix}_Penalty"] = penalties
    return df


def _stage_summary(df: pd.DataFrame, mw_col: str, prefix: str, stage_label: str) -> Optional[dict]:
    """Shared implementation behind step1_summary() / step2_summary(): a
    head-to-head roll-up (blocks both this stage and the official schedule
    covered) -- the official Total_Block_Penalty vs this stage's penalty,
    and mean absolute deviation for each. Returns None if this stage's data
    isn't present."""
    if mw_col not in df.columns:
        return None
    dev_col = f"{prefix}_Deviation_MW"
    covered = df[df[dev_col].notna()] if dev_col in df.columns \
        else df[df[mw_col].notna() & df["Actual_MW"].notna()]
    n = len(covered)
    if n == 0:
        return None
    return {
        "Blocks compared": n,
        "Official Total Penalty (Rs, same blocks)": float(covered["Total_Block_Penalty"].sum()),
        f"{stage_label} Total Penalty (Rs)": float(covered[f"{prefix}_Penalty"].sum()),
        "Official Mean Abs Deviation (MW)": float(covered["Deviation_MW"].abs().mean()),
        f"{stage_label} Mean Abs Deviation (MW)": float(covered[dev_col].abs().mean()),
    }


def evaluate_step1(df: pd.DataFrame, plant: PlantConfig) -> pd.DataFrame:
    """Step 1 (raw meter-base forecast) comparison columns -- see
    _evaluate_comparison_stage(). A no-op if "Step1_MW" isn't present."""
    return _evaluate_comparison_stage(df, plant, "Step1_MW", "Step1")


def step1_summary(df: pd.DataFrame) -> Optional[dict]:
    """Step 1 vs official head-to-head roll-up -- see _stage_summary()."""
    return _stage_summary(df, "Step1_MW", "Step1", "Step 1 (base forecast)")


def evaluate_step2(df: pd.DataFrame, plant: PlantConfig) -> pd.DataFrame:
    """Step 2 (weather-adjusted forecast) comparison columns -- see
    _evaluate_comparison_stage(). A no-op if "Step2_MW" isn't present (e.g.
    when the file only has Step 1 + Step 2 and Step 2 is itself the
    official schedule, so there's nothing to compare it against)."""
    return _evaluate_comparison_stage(df, plant, "Step2_MW", "Step2")


def step2_summary(df: pd.DataFrame) -> Optional[dict]:
    """Step 2 vs official head-to-head roll-up -- see _stage_summary()."""
    return _stage_summary(df, "Step2_MW", "Step2", "Step 2 (weather-adjusted)")


def compute_daily_status(df: pd.DataFrame) -> str:
    """Roll up every block's Status into one day-level label:

        No calculated blocks       -> "Pending"
        Some blocks calculated     -> "Partially Calculated"
        All calculated, total = 0  -> "Zero Penalty"
        All calculated, total > 0  -> "Calculated"
    """
    if "Status" not in df.columns or len(df) == 0:
        return "Pending"
    n_total = len(df)
    n_calc = int((df["Status"] == "Calculated").sum())
    if n_calc == 0:
        return "Pending"
    if n_calc < n_total:
        return "Partially Calculated"
    total_penalty = float(df["Total_Block_Penalty"].sum())
    return "Zero Penalty" if total_penalty == 0 else "Calculated"


def build_summary(df: pd.DataFrame, plant: PlantConfig) -> dict:
    """Compute the KPI summary shown at the top of the dashboard and in the
    Excel 'Summary' sheet.

    "Total Blocks" is the full day (e.g. all 96 blocks), regardless of how
    many actually had matching data. "Blocks Calculated" / "Blocks Pending"
    split that total by Status. Every aggregate below (max/min/mean/sum)
    automatically ignores Pending (NaN) blocks -- pandas skips NaN in these
    by default -- so a missing block is never treated as a zero deviation or
    a zero penalty.
    """
    block_hours = plant.block_hours
    total_blocks = len(df)
    has_status = "Status" in df.columns
    blocks_calculated = int((df["Status"] == "Calculated").sum()) if has_status else total_blocks
    blocks_pending = int((df["Status"] == "Pending").sum()) if has_status else 0

    total_scheduled_energy_mwh = float((df["Scheduled_MW"] * block_hours).sum())
    total_actual_energy_mwh = float((df["Actual_MW"] * block_hours).sum())

    max_pos_dev = float(df["Deviation_MW"].max()) if blocks_calculated else 0.0
    max_neg_dev = float(df["Deviation_MW"].min()) if blocks_calculated else 0.0
    avg_dev = float(df["Deviation_MW"].mean()) if blocks_calculated else 0.0

    total_penalty = float(df["Total_Block_Penalty"].sum())
    avg_penalty_per_block = float(df["Total_Block_Penalty"].mean()) if blocks_calculated else 0.0
    total_ppa_amount = float(df["PPA_Amount"].sum()) if "PPA_Amount" in df.columns else 0.0

    if blocks_calculated and df["Total_Block_Penalty"].notna().any() and df["Total_Block_Penalty"].max() > 0:
        max_pen_idx = df["Total_Block_Penalty"].idxmax()
        max_penalty_block = {
            "Block": int(df.loc[max_pen_idx, "Block"]),
            "Time_Label": df.loc[max_pen_idx, "Time_Label"],
            "Penalty": float(df.loc[max_pen_idx, "Total_Block_Penalty"]),
        }
    else:
        max_penalty_block = {"Block": None, "Time_Label": None, "Penalty": 0.0}

    over_blocks = int((df["Deviation_Type"] == "Over Generation").sum())
    under_blocks = int((df["Deviation_Type"] == "Under Generation").sum())
    exact_blocks = int((df["Deviation_Type"] == "No Deviation").sum())

    return {
        "Plant Name": plant.name,
        "Installed Capacity (MW)": plant.installed_capacity_mw,
        "Total Blocks": total_blocks,
        "Blocks Calculated": blocks_calculated,
        "Blocks Pending": blocks_pending,
        "Daily Status": compute_daily_status(df),
        "Total Scheduled Energy (MWh)": total_scheduled_energy_mwh,
        "Total Actual Energy (MWh)": total_actual_energy_mwh,
        "Energy Deviation (MWh)": total_actual_energy_mwh - total_scheduled_energy_mwh,
        "Maximum Positive Deviation (MW)": max_pos_dev,
        "Maximum Negative Deviation (MW)": max_neg_dev,
        "Average Deviation (MW)": avg_dev,
        "Over Generation Blocks": over_blocks,
        "Under Generation Blocks": under_blocks,
        "No Deviation Blocks": exact_blocks,
        "Maximum Penalty Block": max_penalty_block,
        "Total DSM Penalty (Rs)": total_penalty,
        "Average Penalty per Block (Rs)": avg_penalty_per_block,
        "Total PPA Amount (Rs)": total_ppa_amount,
    }


def day_summary_metrics(df: pd.DataFrame, plant: PlantConfig) -> list:
    """The 'DAY SUMMARY -- ACCURACY AND DSM PENALTY' block shown in the
    Excel report's main sheet, as a single source of truth so the Streamlit
    app's report-preview tab and the Excel export never drift apart.

    Returns a list of dicts: {label, value, kind} where kind is one of
    'int' | 'mw' | 'mw_signed' | 'pct' | 'currency'.
    """
    has_status = "Status" in df.columns
    calc_df = df[df["Status"] == "Calculated"] if has_status else df
    n = len(calc_df)
    n_total = len(df)
    n_pending = n_total - n if has_status else 0
    dev = calc_df["Deviation_MW"].to_numpy() if n else np.array([])
    pen = calc_df["Total_Block_Penalty"].to_numpy() if n else np.array([])
    cap = plant.installed_capacity_mw
    total_ppa = float(df["PPA_Amount"].sum()) if "PPA_Amount" in df.columns else 0.0

    return [
        {"label": "Blocks with a real meter reading", "value": float(n), "kind": "int"},
        {"label": "Blocks Pending (schedule and/or meter missing)",
         "value": float(n_pending), "kind": "int"},
        {"label": "Daily Status", "value": compute_daily_status(df), "kind": "text"},
        {"label": "Total scheduled (MW, scored blocks)",
         "value": float(calc_df["Scheduled_MW"].sum()) if n else 0.0, "kind": "mw"},
        {"label": "Total actual (MW)", "value": float(calc_df["Actual_MW"].sum()) if n else 0.0, "kind": "mw"},
        {"label": "Total deviation — actual minus scheduled (MW)",
         "value": float(dev.sum()) if n else 0.0, "kind": "mw_signed"},
        {"label": "Mean absolute deviation (MW)",
         "value": float(np.abs(dev).mean()) if n else 0.0, "kind": "mw"},
        {"label": "Max absolute deviation (MW)",
         "value": float(np.abs(dev).max()) if n else 0.0, "kind": "mw"},
        {"label": "Blocks over 0.5 MW deviation (RED)",
         "value": float((np.abs(dev) > 0.5).sum()) if n else 0.0, "kind": "int"},
        {"label": "Blocks within 0.5 MW deviation (BLUE)",
         "value": float((np.abs(dev) <= 0.5).sum()) if n else 0.0, "kind": "int"},
        {"label": "Mean absolute deviation (% of capacity)",
         "value": float(np.abs(dev).mean() / cap * 100) if n else 0.0, "kind": "pct"},
        {"label": "Blocks that incurred a penalty",
         "value": float((pen > 0).sum()) if n else 0.0, "kind": "int"},
        {"label": "Worst single-block penalty (Rs)",
         "value": float(pen.max()) if n else 0.0, "kind": "currency"},
        {"label": "TOTAL DSM PENALTY FOR THE DAY (Rs)",
         "value": float(pen.sum()) if n else 0.0, "kind": "currency"},
        {"label": "Total PPA Amount (Rs, reference only)",
         "value": total_ppa, "kind": "currency"},
    ]
