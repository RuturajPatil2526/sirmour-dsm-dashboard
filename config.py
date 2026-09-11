"""
config.py
=========
Central configuration for the AI Schedule Evaluation Dashboard.

Everything that is "plant specific" lives here so the app can be extended to
additional plants (different capacity, different DSM slab structure) without
touching any calculation, UI, or export code.

To add a new plant: add a new entry to PLANT_CONFIGS with its own capacity,
block size and DSM_SLABS list. The rest of the application will pick it up
automatically via the plant selector in the sidebar.
"""

from dataclasses import dataclass, field
from typing import List, Optional, Dict


@dataclass
class DSMSlab:
    """A single DSM penalty slab / band.

    lower / upper are expressed in percent of installed capacity.
    upper = None means "no upper bound" (i.e. the final, open-ended slab).
    rate is in Rs per kWh.
    """
    lower: float
    upper: Optional[float]
    rate: float
    label: str = ""

    def width(self, cap_at: Optional[float] = None) -> float:
        """Width of the slab in percentage points. If the slab is open-ended
        (upper is None) and cap_at is provided, the width is measured up to
        cap_at (used when allocating a specific block's deviation)."""
        upper = self.upper if self.upper is not None else cap_at
        if upper is None:
            return 0.0
        return max(0.0, upper - self.lower)


@dataclass
class PlantConfig:
    name: str
    installed_capacity_mw: float
    block_minutes: int = 15
    dsm_slabs: List[DSMSlab] = field(default_factory=list)
    currency_symbol: str = "₹"  # ₹
    # PPA (Power Purchase Agreement) rate, Rs/kWh -- used ONLY to display a
    # reference "PPA Amount" per block (scheduled energy x this rate). It
    # never affects the DSM penalty calculation itself.
    ppa_rate: float = 0.0

    @property
    def block_hours(self) -> float:
        return self.block_minutes / 60.0

    @property
    def blocks_per_day(self) -> int:
        return int(24 * 60 / self.block_minutes)


# ---------------------------------------------------------------------------
# Default DSM slab structure (as supplied by the business team for SIRMOUR)
# ---------------------------------------------------------------------------
DEFAULT_SLABS = [
    DSMSlab(0, 10, 0.00, "0% - 10%"),
    DSMSlab(10, 15, 0.50, "10% - 15%"),
    DSMSlab(15, 20, 0.75, "15% - 20%"),
    DSMSlab(20, None, 1.00, "Above 20%"),
]

# ---------------------------------------------------------------------------
# Plant registry -- add new plants here
# ---------------------------------------------------------------------------
PLANT_CONFIGS: Dict[str, PlantConfig] = {
    "SIRMOUR": PlantConfig(
        name="SIRMOUR",
        installed_capacity_mw=5.1,
        block_minutes=15,
        dsm_slabs=DEFAULT_SLABS,
        ppa_rate=2.94,
    ),
}

DEFAULT_PLANT = "SIRMOUR"


# ---------------------------------------------------------------------------
# Column alias maps used by utils.py to auto-detect columns in uploaded
# files. Real-world exports rarely match a single fixed schema, so every
# category below lists the many header spellings we should recognise.
# Matching is case-insensitive and ignores surrounding whitespace / units.
# ---------------------------------------------------------------------------
COLUMN_ALIASES = {
    "date": ["date"],
    "block": [
        "block", "block number", "block no", "block_no", "blockno",
        "time block", "timeblock", "sl no", "s.no", "block #",
    ],
    "timestamp": [
        "time", "timestamp", "datetime", "date time", "date_time",
        "generated at interval",
    ],
    "mw_predicted": [
        "predicted_mw", "predicted (mw)", "predicted mw", "predicted",
        "ai schedule (mw)", "ai schedule mw", "schedule_mw",
        "scheduled_mw", "scheduled (mw)", "scheduled mw", "forecast_mw",
        "forecast (mw)",
    ],
    # New 2-step AI Schedule format: "Step 1 Meter Base Forecast MW" (the
    # raw meter-based forecast, before any weather adjustment) and
    # "Step 2 Weather Adjustment MW" (the final, weather-adjusted forecast
    # that is actually submitted as the schedule). When both are present,
    # Step 2 becomes the official Scheduled_MW and Step 1 is kept alongside
    # as its own, separately-penalised comparison column (see utils.py /
    # calculator.evaluate_step1).
    "mw_step1": [
        "step 1 meter base forecast mw", "step 1 meter base forecast (mw)",
        "step1 meter base forecast mw", "step 1 meter base forecast",
        "meter base forecast mw", "meter base forecast (mw)",
        "step 1 mw", "step1 mw", "step 1 (mw)", "step1_mw",
        "base forecast mw", "base forecast (mw)",
    ],
    "mw_step2": [
        "step 2 weather adjustment mw", "step 2 weather adjustment (mw)",
        "step2 weather adjustment mw", "step 2 weather adjustment",
        "weather adjustment mw", "weather adjustment (mw)",
        "step 2 mw", "step2 mw", "step 2 (mw)", "step2_mw",
        "weather adjusted mw", "weather adjusted (mw)",
    ],
    "kw_predicted": [
        "predicted_kw", "predicted (kw)", "predicted kw",
        "ai schedule (kw)", "schedule_kw", "scheduled_kw", "forecast_kw",
    ],
    "mw_actual": [
        "actual_mw", "actual (mw)", "actual mw", "actual",
        "meter_mw", "meter (mw)",
    ],
    "kw_actual": [
        "actual_kw", "actual (kw)", "actual kw", "meter_kw",
        "active power (kw)", "active power(kw)", "activepower(kw)",
        "power (kw)", "power(kw)",
    ],
    "generated_at": [
        "generated at interval", "generated at", "scheduled at",
        "forecast generated at", "forecast time", "generation time",
    ],
}

# Rows in the 96-block day, used as a lookup for human-readable time labels
# e.g. block 1 -> "00:00-00:15"
def block_time_label(block_number: int, block_minutes: int = 15) -> str:
    total_minutes_start = (block_number - 1) * block_minutes
    total_minutes_end = block_number * block_minutes
    start_h, start_m = divmod(total_minutes_start % (24 * 60), 60)
    end_minutes = total_minutes_end % (24 * 60)
    if end_minutes == 0 and total_minutes_end != 0:
        end_h, end_m = 24, 0
        end_label = "24:00"
    else:
        end_h, end_m = divmod(end_minutes, 60)
        end_label = f"{end_h:02d}:{end_m:02d}"
    return f"{start_h:02d}:{start_m:02d}-{end_label}"
