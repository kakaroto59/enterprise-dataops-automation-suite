"""Unit tests for 05_operational_data_pipeline (ETL, Dual-Stream & YAML Linter)."""

from __future__ import annotations

import importlib
from pathlib import Path
import sys

import pandas as pd

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

op = importlib.import_module("src.05_operational_data_pipeline.pipeline")
linter = importlib.import_module("src.05_operational_data_pipeline.yaml_linter")


def test_dual_stream_isolation_ignores_system_events() -> None:
    logs = pd.Series(
        [
            (
                "[[EVENTS]]\n"
                "2026-11-23 | SYSTEM: Priority changed to 2\n"
                "2026-11-23 | FLAGGED_ATTENTION: Urgent check\n"
                "[[INTERACTIONS]]\n"
                "2026-11-10 | Team: Initial filing\n"
                "2026-11-20 | Lead: Latest executive interaction"
            ),
            "[[EVENTS]]\n2026-11-22 | SYSTEM: Created record",
        ]
    )
    parsed = op.parse_dual_stream_history(logs)
    assert parsed.loc[0, "EVENT_COUNT"] == 2
    assert parsed.loc[0, "INTERACTION_COUNT"] == 2
    assert parsed.loc[0, "LATEST_INTERACTION"] == "2026-11-20 | Lead: Latest executive interaction"
    assert pd.Timestamp(parsed.loc[0, "LAST_INTERACTION_DATE"]) == pd.Timestamp("2026-11-20")

    # Record with only [[EVENTS]] has empty LATEST_INTERACTION
    assert parsed.loc[1, "EVENT_COUNT"] == 1
    assert parsed.loc[1, "INTERACTION_COUNT"] == 0
    assert parsed.loc[1, "LATEST_INTERACTION"] == ""


def test_single_table_etl_kpis_matrix_and_html_export(tmp_path: Path) -> None:
    raw = pd.DataFrame(
        [
            {
                "CASE_NUMBER": "CASE-1; CASE-2",
                "THEME": "Topic A",
                "PRIORITY": "1.0",
                "DIVISION_UNIT": "DIV_ALPHA - Full Name",
                "CLASSIFICATION": "BOARD",
                "ANALYSIS_STATUS": "In Progress",
                "EXTERNAL_ATTENTION_FLAG": True,
                "INTERNAL_HISTORY": "[[INTERACTIONS]]\n2026-11-21 | Recent active note",
            },
            {
                "CASE_NUMBER": "CASE-3",
                "THEME": "Topic B",
                "PRIORITY": "2",
                "DIVISION_UNIT": "DIV_BETA",
                "CLASSIFICATION": "DIVISION",
                "ANALYSIS_STATUS": "In Progress",
                "EXTERNAL_ATTENTION_FLAG": False,
                "INTERNAL_HISTORY": "[[INTERACTIONS]]\n2026-11-01 | Stale over 10d note",
            },
            {
                "CASE_NUMBER": "CASE-4",
                "THEME": "Topic C",
                "PRIORITY": "3",
                "DIVISION_UNIT": "DIV_BETA",
                "CLASSIFICATION": "DIVISION",
                "ANALYSIS_STATUS": "Closed",
                "EXTERNAL_ATTENTION_FLAG": False,
                "INTERNAL_HISTORY": "[[INTERACTIONS]]\n2026-10-01 | Closed old case",
            },
        ]
    )

    sanitized = op.sanitize_single_table_processes(raw)
    assert list(sanitized["CASE_NUMBER"]) == ["CASE-1", "CASE-2", "CASE-3", "CASE-4"]
    assert list(sanitized["DIVISION_UNIT"]) == ["DIV_ALPHA", "DIV_ALPHA", "DIV_BETA", "DIV_BETA"]
    assert bool(sanitized.loc[0, "ESCALATED_TO_BOARD"]) is True

    kpis, matrix_df = op.compute_executive_kpis_and_matrix(sanitized, reference_date="2026-11-23")
    assert kpis["total_monitored"] == 4
    assert kpis["in_progress_count"] == 3
    assert kpis["closed_count"] == 1
    assert kpis["high_priority_p1_count"] == 2
    # CASE-3 is stale (22d), CASE-4 is Closed so excluded!
    assert kpis["stale_10d_count"] == 1
    assert kpis["external_attention_requests"] == 2

    # Matrix strictly has DIVISION_UNIT, P1, P2, P3, TOTAL (no Unassigned column)
    assert list(matrix_df.columns) == ["DIVISION_UNIT", "P1", "P2", "P3", "TOTAL"]
    alpha_row = matrix_df[matrix_df["DIVISION_UNIT"] == "DIV_ALPHA"].iloc[0]
    assert int(alpha_row["P1"]) == 2 and int(alpha_row["TOTAL"]) == 2

    html_out = tmp_path / "hub.html"
    json_out = tmp_path / "cases.json"
    op.export_interactive_html_hub(sanitized, kpis, matrix_df, html_out, json_out)
    html_text = html_out.read_text(encoding="utf-8")
    assert "Enterprise Operations &amp; Compliance Intelligence Hub" in html_text
    assert "[Closed]" in html_text
    assert json_out.exists()


def test_declarative_yaml_linter_valid_and_invalid(tmp_path: Path) -> None:
    valid_yaml = tmp_path / "valid.yaml"
    valid_yaml.write_text(
        """- conRoot:
    Control: GroupContainer@1.5.0
    Properties:
      Fill: =ColorValue("#FFFFFF")
    Children:
      - lblOk:
          Control: Label@2.5.1
          Properties:
            Text: ="Valid Label"
""",
        encoding="utf-8",
    )
    ok_res = linter.lint_powerapps_yaml_files([valid_yaml])
    assert ok_res["passed"] is True
    assert ok_res["total_controls"] == 2

    bad_yaml = tmp_path / "broken.yaml"
    bad_yaml.write_text(
        """- conRoot:
    Control: UnknownWidget@9.9.9
    Properties:
      Fill: ColorValue("#FFF"
    Children:
      - conRoot:
          Control: Label@2.5.1
          Properties:
            Text: ="Duplicate"
""",
        encoding="utf-8",
    )
    bad_res = linter.lint_powerapps_yaml_files([bad_yaml])
    assert bad_res["passed"] is False
    joined = "\n".join(bad_res["errors"])
    assert "unverified control type" in joined
    assert "missing '=' formula prefix" in joined
    assert "unbalanced parentheses" in joined
    assert "Duplicate control identifiers" in joined
