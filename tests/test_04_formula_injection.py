"""Unit tests for 04_dynamic_formula_injection & repository sanitization."""

from __future__ import annotations

import base64
import importlib
from pathlib import Path
import sys

import openpyxl
import pandas as pd

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

sync = importlib.import_module("src.04_dynamic_formula_injection.synchronizer")


def test_regional_filter_and_hash_join() -> None:
    schedule = pd.DataFrame(
        [
            {"REGIONAL": "reg_central", "INTERVENTION_ID": "INT-1"},
            {"REGIONAL": "REG_SOUTH", "INTERVENTION_ID": "INT-2"},
            {"REGIONAL": "REG_NORTH", "INTERVENTION_ID": "INT-3"},
        ]
    )
    solutions = pd.DataFrame(
        [
            {"INTERVENTION_ID": "INT-1", "PLAN": "A"},
            {"INTERVENTION_ID": "INT-2", "PLAN": "B"},
            {"INTERVENTION_ID": "INT-3", "PLAN": "C"},
        ]
    )
    pn_filt = sync.filter_schedule_by_region(schedule)
    assert list(pn_filt["INTERVENTION_ID"]) == ["INT-1", "INT-2"]

    sol_filt = sync.hash_join_solutions(pn_filt, solutions)
    assert list(sol_filt["INTERVENTION_ID"]) == ["INT-1", "INT-2"]


def test_synchronize_workbooks_injects_exact_formula_coordinates(tmp_path: Path) -> None:
    src = tmp_path / "src_rollout.xlsx"
    out = tmp_path / "out_rollout.xlsx"

    pn_df = pd.DataFrame(
        [
            {"REGIONAL": "REG_CENTRAL", "SITE_ID": "S1", "EXTRA": "X", "KEY_REF": "K1", "INTERVENTION_ID": "INT-10"},
            {"REGIONAL": "REG_SOUTH", "SITE_ID": "S2", "EXTRA": "Y", "KEY_REF": "K2", "INTERVENTION_ID": "INT-20"},
            {"REGIONAL": "REG_WEST", "SITE_ID": "S3", "EXTRA": "Z", "KEY_REF": "K3", "INTERVENTION_ID": "INT-30"},
        ]
    )
    sol_df = pd.DataFrame(
        [
            {"INTERVENTION_ID": "INT-10", "BAND": "LOW"},
            {"INTERVENTION_ID": "INT-30", "BAND": "HIGH"},
        ]
    )
    with pd.ExcelWriter(src, engine="openpyxl") as writer:
        pn_df.to_excel(writer, sheet_name="PN_SCHEDULE", index=False)
        sol_df.to_excel(writer, sheet_name="SOLUTIONS", index=False)

    stats = sync.synchronize_workbooks(src, out)
    assert stats["pn_rows"] == 2
    assert stats["sol_rows"] == 1

    wb = openpyxl.load_workbook(out)
    ws_pn = wb["PN_REG_SYNC"]
    ws_sol = wb["SOL_REG_SYNC"]

    # 5 data cols + 6 injected control formula cols = 11 cols on PN_REG_SYNC
    assert ws_pn.cell(row=1, column=6).value == "CTRL_CENTRAL"
    assert ws_pn.cell(row=2, column=6).value == "=COUNTIF('CTRL_CENTRAL_BASE'!AF:AF,D2)"
    assert ws_pn.cell(row=3, column=11).value == "=COUNTIF('PRELIM_SOUTH'!Q:Q,D3)"

    # SOL_REG_SYNC: formulas in A..E, payload starting at F
    assert ws_sol["A2"].value == "=XLOOKUP(LEFT(F2,7),'STA_CENTRAL'!B:B,'STA_CENTRAL'!AB:AB)"
    assert ws_sol["D2"].value == "=COUNTIF('PN_REG_SYNC'!E:E,F2)"
    assert ws_sol["E2"].value == "=BZ2"
    assert ws_sol["F1"].value == "INTERVENTION_ID"
    assert ws_sol["F2"].value == "INT-10"
    wb.close()


def test_lgpd_nda_sanitization_and_no_row_loops() -> None:
    encoded_terms = [
        "aHVhd2Vp",
        "Y2xhcm8=",
        "bmV4dGVs",
        "YW5hdGVs",
        "dml2bw==",
        "dGVsZWZvbmljYQ==",
        "Ym9sb21ldHJv",
        "cmVsZ292",
        "Y29uZXhpcw==",
        "ZmVsaXBlLmJhcnJldG8=",
        "am9hby5mb2xpdmVpcmE=",
    ]
    forbidden = [base64.b64decode(t).decode("utf-8") for t in encoded_terms]

    scan_roots = [
        ROOT_DIR / "README.md",
        ROOT_DIR / "src",
        ROOT_DIR / "docs",
        ROOT_DIR / "data" / "mock",
        ROOT_DIR / ".env.example",
    ]
    for root in scan_roots:
        files = [root] if root.is_file() else list(root.rglob("*"))
        for file_path in files:
            if file_path.is_file() and file_path.suffix in {".py", ".md", ".csv", ".txt", ".example", ".yaml"}:
                content = file_path.read_text(encoding="utf-8", errors="ignore").lower()
                for term in forbidden:
                    assert term not in content, f"Confidential term leaked in {file_path}"

    # Verify zero DataFrame row-level iteration anti-patterns in src/
    for py_file in (ROOT_DIR / "src").rglob("*.py"):
        code = py_file.read_text(encoding="utf-8")
        assert ".iterrows(" not in code, f"Found .iterrows() in {py_file}"
        assert ".apply(" not in code, f"Found row-wise .apply() in {py_file}"
