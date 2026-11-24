"""Synthetic, LGPD/NDA-safe mock dataset generator & end-to-end smoke runner."""

from __future__ import annotations

import importlib
from pathlib import Path
import sys

import openpyxl
import pandas as pd

ROOT_DIR = Path(__file__).resolve().parents[2]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

p1 = importlib.import_module("src.01_in_memory_etl_dashboard.pipeline")
p2 = importlib.import_module("src.02_rules_engine_auditor.auditor")
p3 = importlib.import_module("src.03_vectorized_logic_validator.validator")
p4 = importlib.import_module("src.04_dynamic_formula_injection.synchronizer")
p5 = importlib.import_module("src.05_operational_data_pipeline")

MOCK_DIR = ROOT_DIR / "data" / "mock"
OUTPUT_DIR = ROOT_DIR / "data" / "output"


def create_project_01_mocks() -> None:
    """Generate synthetic daily semicolon logs, master rollout schedule, and site filter."""
    new_rows = pd.DataFrame(
        [
            {"STATION": "SITE-1001", "ACTIVITY": "ACT-101", "STAGE_NAME": "21 - Logical Acceptance", "DAYS_IN_STAGE": 3.5},
            {"STATION": "SITE-1001", "ACTIVITY": "ACT-101", "STAGE_NAME": "22 - Traffic Release", "DAYS_IN_STAGE": 2.0},
            {"STATION": "SITE-1001", "ACTIVITY": "ACT-101", "STAGE_NAME": "23 - Final Site Acceptance", "DAYS_IN_STAGE": 8.0},
            {"STATION": "SITE-1002", "ACTIVITY": "ACT-102", "STAGE_NAME": "21 - Logical Acceptance", "DAYS_IN_STAGE": 4.5},
            {"STATION": "", "ACTIVITY": "SITE-1003", "STAGE_NAME": "22 - Traffic Release", "DAYS_IN_STAGE": 1.5},
            {"STATION": "SITE-9999", "ACTIVITY": "ACT-999", "STAGE_NAME": "21 - Logical Acceptance", "DAYS_IN_STAGE": 10.0},
        ]
    )
    reuse_rows = pd.DataFrame(
        [
            {"STATION": "IGNORED-1", "ACTIVITY": "SITE-1001", "STAGE_NAME": "22 - Execute Field Test", "DAYS_IN_STAGE": 2.0},
            {"STATION": "IGNORED-2", "ACTIVITY": "SITE-1001", "STAGE_NAME": "24 - Submit Acceptance CRQ", "DAYS_IN_STAGE": 3.0},
            {"STATION": "IGNORED-3", "ACTIVITY": "SITE-1002", "STAGE_NAME": "26 - Traffic Release", "DAYS_IN_STAGE": 1.0},
            {"STATION": "IGNORED-4", "ACTIVITY": "SITE-1002", "STAGE_NAME": "27 - Final Element Acceptance", "DAYS_IN_STAGE": 12.0},
        ]
    )
    new_csv = MOCK_DIR / "23112026-new_flow.csv"
    reuse_csv = MOCK_DIR / "23112026-reuse_flow.csv"
    new_rows.to_csv(new_csv, sep=";", index=False, encoding="utf-8")
    reuse_rows.to_csv(reuse_csv, sep=";", index=False, encoding="utf-8")

    schedule_df = pd.DataFrame(
        [
            {"SITE_ID": "SITE-1001", "REGIONAL": "REG_CENTRAL", "PRIORITY_TIER": "P1", "VENDOR_TEAM": "TEAM_A"},
            {"SITE_ID": "SITE-1002", "REGIONAL": "REG_SOUTH", "PRIORITY_TIER": "P1", "VENDOR_TEAM": "TEAM_B"},
            {"SITE_ID": "SITE-1003", "REGIONAL": "REG_CENTRAL", "PRIORITY_TIER": "P2", "VENDOR_TEAM": "TEAM_A"},
            {"SITE_ID": "SITE-9999", "REGIONAL": "REG_NORTH", "PRIORITY_TIER": "P3", "VENDOR_TEAM": "TEAM_C"},
        ]
    )
    schedule_xlsx = MOCK_DIR / "master_rollout_schedule.xlsx"
    schedule_df.to_excel(schedule_xlsx, index=False)

    target_sites = pd.DataFrame({"SITE_ID": ["SITE-1001", "SITE-1002", "SITE-1003", "SITE-9999"]})
    sites_csv = MOCK_DIR / "target_sites.csv"
    target_sites.to_csv(sites_csv, index=False)

    p1.run_etl_pipeline(
        log_entries=[(new_csv, "New"), (reuse_csv, "Reuse")],
        schedule_path=schedule_xlsx,
        output_path=OUTPUT_DIR / "consolidated_sla_dashboard.xlsx",
        sites_file=sites_csv,
    )


def create_project_02_mocks() -> None:
    """Generate synthetic compliance audit workbook with row-2 header and duplicate column."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "COMPLIANCE_BASE"

    ws.append(["SYSTEM EXPORT METADATA - SYNTHETIC MOCK DATASET"])
    # Header row at index 1 (Excel row 2); index 5 is official OBSERVATION, index 10 is duplicate OBSERVATION
    ws.append(
        [
            "SITE_ID",
            "RECORD_KEY",
            "SI_STATUS",
            "ROLLOUT_STATUS",
            "STATION_STATE",
            "OBSERVATION",
            "AUDIT_OWNER_STATUS",
            "FLOW_CATEGORY",
            "IW_STAGE",
            "SOLUTION",
            "OBSERVATION",
        ]
    )
    rows = [
        # SITE-2001: 2 sectors, both pass -> PROMOTE_TO_CERTIFIED
        ["SITE-2001", "SITE-2001_SEC1", "OK", "OK", "Mobilized", "Ready sector 1", "PENDING", "New", "20 - Commissioning", "LOW/MID", "Trailing note 1"],
        ["SITE-2001", "SITE-2001_SEC2", "OK", "OK", "In_Formalization", "Ready sector 2", "PENDING", "Remote_Sector_Reuse", "12 - Field Validation", "MID", "Trailing note 2"],
        # SITE-2002: 2 sectors, 1 fails -> atomic site rejection (no promotion)
        ["SITE-2002", "SITE-2002_SEC1", "OK", "OK", "Mobilized", "Sector 1 OK", "PENDING", "New", "19 - Integration", "LOW", "Trailing note 3"],
        ["SITE-2002", "SITE-2002_SEC2", "OK", "OK", "Mobilized", "Sector 2 Out of bounds", "PENDING", "New", "10 - Early Civil", "LOW", "Trailing note 4"],
        # SITE-2003: Already CERTIFIED, 1 sector regressed -> AUDIT_REGRESSION_ALERT for whole site
        ["SITE-2003", "SITE-2003_SEC1", "OK", "OK", "Mobilized", "Sector 1 OK", "CERTIFIED", "Remote_Sector_SN", "14 - Approval", "HIGH", "Trailing note 5"],
        ["SITE-2003", "SITE-2003_SEC2", "OK", "OK", "Mobilized", "Sector 2 Regressed", "CERTIFIED", "Reuse", "25 - Blocked", "HIGH", "Trailing note 6"],
    ]
    for r in rows:
        ws.append(r)

    wb_path = MOCK_DIR / "compliance_audit_workbook.xlsx"
    wb.save(wb_path)
    wb.close()

    p2.run_audit_file(
        input_path=wb_path,
        output_path=OUTPUT_DIR / "compliance_audit_report.xlsx",
    )


def create_project_03_mocks() -> None:
    """Generate synthetic multi-tier frequency matrix workbook."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "FREQ_MATRIX"

    ws.append(["SYNTHETIC RF DEPLOYMENT MATRIX"])
    ws.append(["SITE_ID", "TECH_TIER", "SOLUTION", "TIER_1", "TIER_2", "TIER_3", "TIER_4", "BW_DUAL_ACTIVITY", "BW_EXT_ACTIVITY"])
    rows = [
        ["SITE-3001", "TIER_1", "BASE", "BASE_LOW", "", "", "", "-", "-"],
        ["SITE-3002", "TIER_2", "DUAL/LOW", "", "DUALA_DUALB_LOW", "", "", "ALPHA+BETA", "-"],
        ["SITE-3003", "TIER_2", "DUAL", "", "DUALA_LOW", "", "", "ALPHA+BETA", "-"],
        ["SITE-3004", "TIER_3", "LOW/HIGH/MID", "", "", "LOWA_HIGHB_MIDC_MIDEXT", "", "-", "EXT_BW"],
        ["SITE-3005", "TIER_3", "LOW/MID", "", "", "LOWA_MIDC", "", "-", "EXT_BW"],
        ["SITE-3006", "TIER_4", "ULTRA", "", "", "", "ULTRA_HIGH", "-", "-"],
        ["SITE-3007", "TIER_3", "#N/A", "", "", "LOWA", "", "-", "-"],
        ["SITE-3008", "TIER_2", "0", "", "", "", "", "-", "-"],
    ]
    for r in rows:
        ws.append(r)

    wb_path = MOCK_DIR / "frequency_validation_workbook.xlsx"
    wb.save(wb_path)
    wb.close()

    p3.run_validator_file(
        input_path=wb_path,
        output_path=OUTPUT_DIR / "frequency_validation_checked.xlsx",
        target_column_letter="J",
    )


def create_project_04_mocks() -> None:
    """Generate synthetic secondary rollout workbook with PN_SCHEDULE and SOLUTIONS."""
    schedule_df = pd.DataFrame(
        [
            {"REGIONAL": "REG_CENTRAL", "SITE_ID": "SITE-4001", "SITE_SHORT": "S4001", "KEY_REF": "K-4001", "INTERVENTION_ID": "INT-1001", "VENDOR": "VENDOR_X"},
            {"REGIONAL": "REG_SOUTH", "SITE_ID": "SITE-4002", "SITE_SHORT": "S4002", "KEY_REF": "K-4002", "INTERVENTION_ID": "INT-1002", "VENDOR": "VENDOR_Y"},
            {"REGIONAL": "REG_NORTH", "SITE_ID": "SITE-4003", "SITE_SHORT": "S4003", "KEY_REF": "K-4003", "INTERVENTION_ID": "INT-9999", "VENDOR": "VENDOR_Z"},
        ]
    )
    solutions_df = pd.DataFrame(
        [
            {"INTERVENTION_ID": "INT-1001", "BAND_PLAN": "LOW/MID", "STATUS": "APPROVED"},
            {"INTERVENTION_ID": "INT-1002", "BAND_PLAN": "HIGH", "STATUS": "PENDING"},
            {"INTERVENTION_ID": "INT-9999", "BAND_PLAN": "ULTRA", "STATUS": "IGNORED_REGION"},
        ]
    )
    src_path = MOCK_DIR / "secondary_rollout_source.xlsx"
    with pd.ExcelWriter(src_path, engine="openpyxl") as writer:
        schedule_df.to_excel(writer, sheet_name="PN_SCHEDULE", index=False)
        solutions_df.to_excel(writer, sheet_name="SOLUTIONS", index=False)

    p4.synchronize_workbooks(
        source_workbook=src_path,
        output_workbook=OUTPUT_DIR / "synchronized_rollout_control.xlsx",
    )


def create_project_05_mocks() -> None:
    """Generate synthetic single-table operational dataset, sample YAML UI tree, and HTML Hub."""
    tracker_df = pd.DataFrame(
        [
            {
                "CASE_NUMBER": "CASE-2026-1001",
                "THEME": "Spectrum Licensing Review",
                "SYNTHESIS": "Primary regional spectrum compliance evaluation",
                "PRIORITY": "1",
                "DIVISION_UNIT": "DIV_SPECTRUM - Spectrum Engineering",
                "CLASSIFICATION": "BOARD",
                "ANALYSIS_STATUS": "In Progress",
                "INTERNAL_OWNER": "Analyst Alpha",
                "EXTERNAL_CONTACT": "Director One",
                "EXTERNAL_ATTENTION_FLAG": True,
                "CASE_LINK": "https://portal.example.org/cases/1001",
                "INTERNAL_HISTORY": (
                    "[[EVENTS]]\n"
                    "2026-11-22 | SYSTEM: Priority escalated to 1\n"
                    "2026-11-22 | FLAGGED_ATTENTION: Needs board briefing before Q4 cutoff\n"
                    "[[INTERACTIONS]]\n"
                    "2026-11-10 | Ops Lead: Submitted technical dossier\n"
                    "2026-11-20 | Tech Lead: Final executive brief reviewed"
                ),
            },
            {
                "CASE_NUMBER": "CASE-2026-1002; CASE-2026-1003",
                "THEME": "Rural Connectivity Audit",
                "SYNTHESIS": "Coverage obligation verification across rural districts",
                "PRIORITY": "2.0",
                "DIVISION_UNIT": "DIV_COVERAGE",
                "CLASSIFICATION": "DIVISION",
                "ANALYSIS_STATUS": "In Progress",
                "INTERNAL_OWNER": "Analyst Beta",
                "EXTERNAL_CONTACT": "Director Two",
                "EXTERNAL_ATTENTION_FLAG": False,
                "CASE_LINK": "",
                "INTERNAL_HISTORY": (
                    "[[EVENTS]]\n"
                    "2026-11-19 | SYSTEM: Status updated to In Progress\n"
                    "[[INTERACTIONS]]\n"
                    "2026-11-01 | Ops Team: Initial field validation pending"
                ),
            },
            {
                "CASE_NUMBER": "CASE-2026-1004",
                "THEME": "Legacy Microwave Decommission",
                "SYNTHESIS": "Decommission close-out report",
                "PRIORITY": "3",
                "DIVISION_UNIT": "DIV_INFRA",
                "CLASSIFICATION": "DIVISION",
                "ANALYSIS_STATUS": "Closed",
                "INTERNAL_OWNER": "Analyst Gamma",
                "EXTERNAL_CONTACT": "Director Three",
                "EXTERNAL_ATTENTION_FLAG": False,
                "CASE_LINK": "https://portal.example.org/cases/1004",
                "INTERNAL_HISTORY": (
                    "[[EVENTS]]\n"
                    "2026-11-02 | SYSTEM: Case closed after final sign-off\n"
                    "[[INTERACTIONS]]\n"
                    "2026-10-15 | Ops Team: Decommission sign-off archived"
                ),
            },
        ]
    )
    tracker_xlsx = MOCK_DIR / "single_table_operations_tracker.xlsx"
    tracker_df.to_excel(tracker_xlsx, index=False)

    sample_yaml = """- conMainHub:
    Control: GroupContainer@1.5.0
    Variant: AutoLayout
    Properties:
      Fill: =ColorValue("#F5F6FA")
      Height: =Parent.Height
      Width: =Parent.Width
    Children:
      - lblHeaderTitle:
          Control: Label@2.5.1
          Properties:
            Color: =Color.White
            Text: ="Operations Intelligence Hub"
      - btnPriorityLegend:
          Control: Classic/Button@2.2.0
          Properties:
            OnSelect: =Set(varShowLegend, !varShowLegend)
            Text: ="Priority Legend"
"""
    yaml_path = MOCK_DIR / "sample_declarative_ui.yaml"
    yaml_path.write_text(sample_yaml, encoding="utf-8")

    p5.run_operational_pipeline(
        input_path=tracker_xlsx,
        output_html_path=OUTPUT_DIR / "operations_intelligence_hub.html",
        output_json_path=OUTPUT_DIR / "single_table_cases.json",
        reference_date="2026-11-23",
    )
    lint_res = p5.lint_powerapps_yaml_files([yaml_path])
    assert lint_res["passed"], f"Sample YAML lint failed: {lint_res['errors']}"


def main() -> None:
    MOCK_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    create_project_01_mocks()
    create_project_02_mocks()
    create_project_03_mocks()
    create_project_04_mocks()
    create_project_05_mocks()
    print("Generated all synthetic mock inputs and executed all 5 pipelines successfully.")


if __name__ == "__main__":
    main()
