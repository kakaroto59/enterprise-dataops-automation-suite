"""Unit tests for 03_vectorized_logic_validator."""

from __future__ import annotations

import importlib
from pathlib import Path
import sys

import pandas as pd

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

validator = importlib.import_module("src.03_vectorized_logic_validator.validator")


def _make_row(**overrides: object) -> dict[str, object]:
    base = {
        "TECH_TIER": "TIER_3",
        "SOLUTION": "",
        "TIER_1": "",
        "TIER_2": "",
        "TIER_3": "",
        "TIER_4": "",
        "BW_DUAL_ACTIVITY": "-",
        "BW_EXT_ACTIVITY": "-",
    }
    base.update(overrides)
    return base


def test_priority_and_multitier_matrix() -> None:
    rows = [
        # Priority 1: Check Schedule
        _make_row(SOLUTION="#N/A", TIER_3="LOW"),
        _make_row(SOLUTION="#N/D", TIER_3="LOW"),
        # Priority 2: OK SWAP
        _make_row(SOLUTION=None),
        _make_row(SOLUTION="0.0"),
        _make_row(SOLUTION="NONE"),
        # TIER_1 strict
        _make_row(TECH_TIER="TIER_1", SOLUTION="BASE/AUX", TIER_1="BASE_AUX_EXTRA"),
        _make_row(TECH_TIER="TIER_1", SOLUTION="BASE/AUX", TIER_1="BASE"),
        # TIER_2 dual-carrier rules
        _make_row(TECH_TIER="TIER_2", SOLUTION="DUAL/LOW", TIER_2="DUALA_DUALB_LOW", BW_DUAL_ACTIVITY="ALPHA+BETA"),
        _make_row(TECH_TIER="TIER_2", SOLUTION="DUAL", TIER_2="DUALA", BW_DUAL_ACTIVITY="ALPHA+BETA"),
        _make_row(TECH_TIER="TIER_2", SOLUTION="DUAL", TIER_2="DUALB", BW_DUAL_ACTIVITY="BETA"),
        _make_row(TECH_TIER="TIER_2", SOLUTION="DUAL", TIER_2="DUAL", BW_DUAL_ACTIVITY="-"),
        # TIER_3 alias normalization + EXT_BW
        _make_row(TECH_TIER="TIER_3", SOLUTION="LOW/HIGH/MID", TIER_3="LOWA_HIGHB_MIDC_MIDEXT", BW_EXT_ACTIVITY="EXT_BW"),
        _make_row(TECH_TIER="TIER_3", SOLUTION="LOW/MID", TIER_3="LOWA_MIDC", BW_EXT_ACTIVITY="EXT_BW"),
        # TIER_4 strict
        _make_row(TECH_TIER="TIER_4", SOLUTION="ULTRA", TIER_4="ULTRA"),
    ]
    df = pd.DataFrame(rows)
    # Use a non-contiguous index to verify positional alignment safety
    df.index = [100 + i * 3 for i in range(len(df))]

    statuses = validator.validate_frequency_matrix(df).tolist()
    assert statuses == [
        "Check Schedule",
        "Check Schedule",
        "OK SWAP",
        "OK SWAP",
        "OK SWAP",
        "OK",
        "NOK",
        "OK",
        "NOK",
        "OK",
        "OK",
        "OK",
        "NOK",
        "OK",
    ]


def test_vectorized_batch_scale() -> None:
    batch = pd.DataFrame(
        [_make_row(TECH_TIER="TIER_3", SOLUTION="LOW/HIGH", TIER_3="LOWA_HIGHA") for _ in range(2500)]
        + [_make_row(TECH_TIER="TIER_3", SOLUTION="LOW/HIGH", TIER_3="LOWA") for _ in range(2500)]
    )
    res = validator.validate_frequency_matrix(batch)
    assert (res.iloc[:2500] == "OK").all()
    assert (res.iloc[2500:] == "NOK").all()
