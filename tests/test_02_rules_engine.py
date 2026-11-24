"""Unit tests for 02_rules_engine_auditor."""

from __future__ import annotations

import importlib
from pathlib import Path
import sys

import openpyxl
import pandas as pd

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

auditor = importlib.import_module("src.02_rules_engine_auditor.auditor")


def test_resolve_duplicate_headers_preserves_index_5() -> None:
    raw = ["SITE_ID", "RECORD_KEY", "SI_STATUS", "ROLLOUT_STATUS", "STATION_STATE", "OBSERVATION", "A", "B", "OBSERVATION"]
    resolved = auditor.resolve_duplicate_headers(raw, primary_col_name="OBSERVATION", primary_index=5)
    assert resolved[5] == "OBSERVATION"
    assert resolved[8] == "OBSERVATION_DISCARD_8"


def test_atomic_site_approval_and_regression_detection() -> None:
    df = pd.DataFrame(
        [
            # SITE-1: 2 sectors, both pass -> PROMOTE_TO_CERTIFIED
            {"SITE_ID": "SITE-1", "RECORD_KEY": "R1", "SI_STATUS": "OK", "ROLLOUT_STATUS": "OK", "STATION_STATE": "Mobilized", "OBSERVATION": "Note 1", "AUDIT_OWNER_STATUS": "PENDING", "FLOW_CATEGORY": "New", "IW_STAGE": "18 - Start", "SOLUTION": "LOW"},
            {"SITE_ID": "SITE-1", "RECORD_KEY": "R2", "SI_STATUS": "OK", "ROLLOUT_STATUS": "OK", "STATION_STATE": "In_Formalization", "OBSERVATION": "Note 2", "AUDIT_OWNER_STATUS": "PENDING", "FLOW_CATEGORY": "Remote_Sector_Reuse", "IW_STAGE": "16 - Ready", "SOLUTION": "MID"},
            # SITE-2: 2 sectors, 1 fails -> atomic rollback (must NOT promote R3)
            {"SITE_ID": "SITE-2", "RECORD_KEY": "R3", "SI_STATUS": "OK", "ROLLOUT_STATUS": "OK", "STATION_STATE": "Mobilized", "OBSERVATION": "Note 3", "AUDIT_OWNER_STATUS": "PENDING", "FLOW_CATEGORY": "New", "IW_STAGE": "22 - Pass", "SOLUTION": "LOW"},
            {"SITE_ID": "SITE-2", "RECORD_KEY": "R4", "SI_STATUS": "OK", "ROLLOUT_STATUS": "OK", "STATION_STATE": "Mobilized", "OBSERVATION": "Note 4", "AUDIT_OWNER_STATUS": "PENDING", "FLOW_CATEGORY": "New", "IW_STAGE": "23 - OutOfRange", "SOLUTION": "LOW"},
            # SITE-3: CERTIFIED, 1 sector fails -> AUDIT_REGRESSION_ALERT for both sectors of SITE-3
            {"SITE_ID": "SITE-3", "RECORD_KEY": "R5", "SI_STATUS": "OK", "ROLLOUT_STATUS": "OK", "STATION_STATE": "Mobilized", "OBSERVATION": "Note 5", "AUDIT_OWNER_STATUS": "CERTIFIED", "FLOW_CATEGORY": "Remote_Sector_SN", "IW_STAGE": "7 - MinValid", "SOLUTION": "HIGH"},
            {"SITE_ID": "SITE-3", "RECORD_KEY": "R6", "SI_STATUS": "OK", "ROLLOUT_STATUS": "OK", "STATION_STATE": "Mobilized", "OBSERVATION": "Note 6", "AUDIT_OWNER_STATUS": "CERTIFIED", "FLOW_CATEGORY": "Reuse", "IW_STAGE": "6 - Regressed", "SOLUTION": "HIGH"},
        ]
    )

    result_df, summary = auditor.audit_compliance(df)
    assert summary == {"eligible_rows": 6, "promote_count": 2, "regression_count": 2}
    assert set(result_df["SITE_ID"]) == {"SITE-1", "SITE-3"}
    assert "SITE-2" not in result_df["SITE_ID"].values
    assert set(result_df.loc[result_df["SITE_ID"] == "SITE-1", "SUGGESTED_ACTION"]) == {"PROMOTE_TO_CERTIFIED"}
    assert set(result_df.loc[result_df["SITE_ID"] == "SITE-3", "SUGGESTED_ACTION"]) == {"AUDIT_REGRESSION_ALERT"}


def test_run_audit_file_end_to_end(tmp_path: Path) -> None:
    src = tmp_path / "audit_in.xlsx"
    out = tmp_path / "audit_out.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "COMPLIANCE_BASE"
    ws.append(["Metadata Banner"])
    ws.append(["SITE_ID", "RECORD_KEY", "SI_STATUS", "ROLLOUT_STATUS", "STATION_STATE", "OBSERVATION", "AUDIT_OWNER_STATUS", "FLOW_CATEGORY", "IW_STAGE", "SOLUTION", "OBSERVATION"])
    ws.append(["SITE-90", "RK-90", "OK", "OK", "Mobilized", "Canonical Obs", "PENDING", "New", "20 - Valid", "LOW", "Discard Obs"])
    wb.save(src)
    wb.close()

    res_df, summary = auditor.run_audit_file(src, out)
    assert summary["promote_count"] == 1
    assert res_df["OBSERVATION"].iloc[0] == "Canonical Obs"

    wb_out = openpyxl.load_workbook(out)
    assert "FILTERED_AUDIT_RESULTS" in wb_out.sheetnames
    wb_out.close()
