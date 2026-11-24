"""Unit tests for 01_in_memory_etl_dashboard."""

from __future__ import annotations

from datetime import date
import importlib
from pathlib import Path
import sys

import openpyxl
import pandas as pd

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

etl = importlib.import_module("src.01_in_memory_etl_dashboard.pipeline")


def test_extract_date_from_filename() -> None:
    assert etl.extract_date_from_filename("C:/logs/23112026-reuse.csv") == date(2026, 11, 23)
    assert etl.extract_date_from_filename("prefix-05102026-new_flow.txt") == date(2026, 10, 5)
    assert etl.extract_date_from_filename("no_date_here.csv") is None


def test_consolidate_site_id_resolution_and_deduplication(tmp_path: Path) -> None:
    new_csv = tmp_path / "23112026-new.csv"
    reuse_csv = tmp_path / "23112026-reuse.csv"

    pd.DataFrame(
        [
            {"STATION": "SITE-A", "ACTIVITY": "ACT-1", "STAGE_NAME": "21 - Logical Acceptance", "DAYS_IN_STAGE": 3},
            {"STATION": "", "ACTIVITY": "SITE-B", "STAGE_NAME": "22 - Traffic Release", "DAYS_IN_STAGE": 5},
            # Duplicate snapshot key keeps latest entry
            {"STATION": "", "ACTIVITY": "SITE-B", "STAGE_NAME": "22 - Traffic Release", "DAYS_IN_STAGE": 7},
        ]
    ).to_csv(new_csv, sep=";", index=False)

    pd.DataFrame(
        [
            {"STATION": "IGNORED", "ACTIVITY": "SITE-C", "STAGE_NAME": "22 - Execute Field Test", "DAYS_IN_STAGE": 4},
        ]
    ).to_csv(reuse_csv, sep=";", index=False)

    df = etl.consolidate_raw_logs([(new_csv, "New"), (reuse_csv, "Reuse")])
    assert len(df) == 3
    site_b = df[(df["SITE_ID"] == "SITE-B") & (df["FLOW_TYPE"] == "New")]
    assert len(site_b) == 1
    assert float(site_b["DAYS_IN_STAGE"].iloc[0]) == 7.0
    assert "SITE-C" in df[df["FLOW_TYPE"] == "Reuse"]["SITE_ID"].tolist()


def test_sla_calculation_excludes_terminal_stages_and_exports_charts(tmp_path: Path) -> None:
    raw = pd.DataFrame(
        [
            {"SNAPSHOT_DATE": pd.Timestamp("2026-11-23"), "FLOW_TYPE": "New", "SITE_ID": "SITE-1", "STAGE_NAME": "21 - Logical Acceptance", "DAYS_IN_STAGE": 4.0},
            {"SNAPSHOT_DATE": pd.Timestamp("2026-11-23"), "FLOW_TYPE": "New", "SITE_ID": "SITE-1", "STAGE_NAME": "22 - Traffic Release", "DAYS_IN_STAGE": 2.0},
            {"SNAPSHOT_DATE": pd.Timestamp("2026-11-23"), "FLOW_TYPE": "New", "SITE_ID": "SITE-1", "STAGE_NAME": "23 - Final Site Acceptance", "DAYS_IN_STAGE": 99.0},
            {"SNAPSHOT_DATE": pd.Timestamp("2026-11-23"), "FLOW_TYPE": "Reuse", "SITE_ID": "SITE-1", "STAGE_NAME": "22 - Execute Field Test", "DAYS_IN_STAGE": 5.0},
            {"SNAPSHOT_DATE": pd.Timestamp("2026-11-23"), "FLOW_TYPE": "Reuse", "SITE_ID": "SITE-1", "STAGE_NAME": "27 - Final Element Acceptance", "DAYS_IN_STAGE": 50.0},
            {"SNAPSHOT_DATE": pd.Timestamp("2026-11-23"), "FLOW_TYPE": "New", "SITE_ID": "SITE-OUT", "STAGE_NAME": "21 - Logical Acceptance", "DAYS_IN_STAGE": 10.0},
        ]
    )
    sched = pd.DataFrame(
        [
            {"SITE_ID": "SITE-1", "REGIONAL": "REG_CENTRAL"},
            {"SITE_ID": "SITE-OUT", "REGIONAL": "REG_NORTH"},
        ]
    )
    merged = etl.filter_and_merge_schedule(raw, sched, target_regions=("REG_CENTRAL", "REG_SOUTH"))
    assert "SITE-OUT" not in merged["SITE_ID"].values

    new_df, reuse_df, summary_df, totals_df = etl.compute_stage_slas(merged)
    new_total = float(totals_df.loc[totals_df["FLOW_TYPE"] == "New", "TOTAL_ACTIVE_CYCLE_DAYS"].iloc[0])
    reuse_total = float(totals_df.loc[totals_df["FLOW_TYPE"] == "Reuse", "TOTAL_ACTIVE_CYCLE_DAYS"].iloc[0])

    assert new_total == 6.0  # 4.0 + 2.0 (excludes 99.0 from stage 23)
    assert reuse_total == 5.0  # excludes 50.0 from stage 27

    out_xlsx = tmp_path / "dashboard.xlsx"
    etl.export_dashboard_workbook(new_df, reuse_df, summary_df, totals_df, out_xlsx)

    wb = openpyxl.load_workbook(out_xlsx)
    assert set(wb.sheetnames) == {"SLA_Dashboard", "New_Flow_Details", "Reuse_Flow_Details"}
    assert len(wb["SLA_Dashboard"]._charts) == 2
    wb.close()
