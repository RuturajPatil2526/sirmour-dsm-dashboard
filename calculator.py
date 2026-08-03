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
    dataframe and compute every required DSM column, block by block.
    """
    df = merged_df.copy().reset_index(drop=True)
    capacity = plant.installed_capacity_mw
    block_hours = plant.block_hours
    slabs = plant.dsm_slabs

    df["Deviation_MW"] = df["Actual_MW"] - df["Scheduled_MW"]
    df["Deviation_%"] = (df["Deviation_MW"] / capacity) * 100.0
    df["Abs_Deviation_%"] = df["Deviation_%"].abs()
    df["Block_Energy_kWh"] = df["Deviation_MW"] * block_hours * 1000.0

    energy_cols, penalty_cols = _slab_columns(slabs)
    for c in energy_cols + penalty_cols:
        df[c] = 0.0

    total_penalty = np.zeros(len(df))

    for idx, row in df.iterrows():
        abs_pct = row["Abs_Deviation_%"]
        energy_abs = abs(row["Block_Energy_kWh"])
        energies, penalties, tot = compute_block_penalty(abs_pct, energy_abs, slabs)
        for c, v in zip(energy_cols, energies):
            df.at[idx, c] = v
        for c, v in zip(penalty_cols, penalties):
            df.at[idx, c] = v
        total_penalty[idx] = tot

    df["Total_Block_Penalty"] = total_penalty
    df["Deviation_Type"] = np.where(
        df["Deviation_MW"] > 1e-9, "Over Generation",
        np.where(df["Deviation_MW"] < -1e-9, "Under Generation", "No Deviation"),
    )

    ordered_cols = (
        ["Date", "Block", "Time_Label", "Scheduled_MW", "Actual_MW",
         "Deviation_MW", "Deviation_%", "Abs_Deviation_%", "Block_Energy_kWh"]
        + energy_cols + penalty_cols
        + ["Total_Block_Penalty", "Deviation_Type"]
    )
    if "Generated_At" in df.columns:
        ordered_cols.append("Generated_At")
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

    has_enercast = df["Enercast_MW"].notna()
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
    covered = df[df["Enercast_MW"].notna()]
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


def build_summary(df: pd.DataFrame, plant: PlantConfig) -> dict:
    """Compute the KPI summary shown at the top of the dashboard and in the
    Excel 'Summary' sheet.
    """
    block_hours = plant.block_hours
    total_blocks = len(df)
    total_scheduled_energy_mwh = float((df["Scheduled_MW"] * block_hours).sum())
    total_actual_energy_mwh = float((df["Actual_MW"] * block_hours).sum())

    max_pos_dev = float(df["Deviation_MW"].max()) if total_blocks else 0.0
    max_neg_dev = float(df["Deviation_MW"].min()) if total_blocks else 0.0
    avg_dev = float(df["Deviation_MW"].mean()) if total_blocks else 0.0

    total_penalty = float(df["Total_Block_Penalty"].sum())
    avg_penalty_per_block = float(df["Total_Block_Penalty"].mean()) if total_blocks else 0.0

    if total_blocks and df["Total_Block_Penalty"].max() > 0:
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
    }


def day_summary_metrics(df: pd.DataFrame, plant: PlantConfig) -> list:
    """The 'DAY SUMMARY -- ACCURACY AND DSM PENALTY' block shown in the
    Excel report's main sheet, as a single source of truth so the Streamlit
    app's report-preview tab and the Excel export never drift apart.

    Returns a list of dicts: {label, value, kind} where kind is one of
    'int' | 'mw' | 'mw_signed' | 'pct' | 'currency'.
    """
    n = len(df)
    dev = df["Deviation_MW"].to_numpy() if n else np.array([])
    pen = df["Total_Block_Penalty"].to_numpy() if n else np.array([])
    cap = plant.installed_capacity_mw

    return [
        {"label": "Blocks with a real meter reading", "value": float(n), "kind": "int"},
        {"label": "Total scheduled (MW, scored blocks)",
         "value": float(df["Scheduled_MW"].sum()) if n else 0.0, "kind": "mw"},
        {"label": "Total actual (MW)", "value": float(df["Actual_MW"].sum()) if n else 0.0, "kind": "mw"},
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
    ]
