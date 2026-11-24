"""Dynamic Formula Injection & Data Synchronization Pipeline.

Refactors heavy Excel VBA Dictionary/AutoFilter copy-paste workflows into a
vectorized Python pipeline that filters master rollout schedules by regional
scope, performs an O(N) in-memory hash join against solutions data (eliminating
slow row-deletion loops), and dynamically injects Excel XLOOKUP/COUNTIF control
formulas sized precisely to the synchronized row count.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Sequence

import openpyxl
import pandas as pd

DEFAULT_TARGET_REGIONS: tuple[str, ...] = ("REG_CENTRAL", "REG_SOUTH")

SCHEDULE_CONTROL_HEADERS: tuple[str, ...] = (
    "CTRL_CENTRAL",
    "REGULATORY_CENTRAL",
    "PRELIM_CENTRAL",
    "CTRL_SOUTH",
    "REGULATORY_SOUTH",
    "PRELIM_SOUTH",
)

SOLUTION_FORMULA_HEADERS: tuple[str, ...] = (
    "INFRA_TYPE",
    "CLASSIFICATION",
    "ALT_SOLUTION",
    "CROSS_CHECK",
    "FULL_SCHEDULE_REF",
)


def filter_schedule_by_region(
    schedule_df: pd.DataFrame,
    target_regions: Sequence[str] = DEFAULT_TARGET_REGIONS,
) -> pd.DataFrame:
    """Filter master rollout schedule vectorially by regional scope."""
    work = schedule_df.copy()
    work.columns = [str(c).strip() for c in work.columns]
    if "REGIONAL" not in work.columns:
        raise ValueError("Master schedule is missing required 'REGIONAL' column.")

    region_set = {r.strip().upper() for r in target_regions}
    mask = work["REGIONAL"].fillna("").astype(str).str.strip().str.upper().isin(region_set)
    return work.loc[mask].reset_index(drop=True)


def hash_join_solutions(
    filtered_schedule_df: pd.DataFrame,
    solutions_df: pd.DataFrame,
    schedule_key_col: str = "INTERVENTION_ID",
    solution_key_col: str = "INTERVENTION_ID",
) -> pd.DataFrame:
    """Perform O(N) in-memory hash-set filtering of solutions without row deletion."""
    if filtered_schedule_df.empty or solutions_df.empty:
        return solutions_df.iloc[0:0].copy()

    sched = filtered_schedule_df
    sol = solutions_df.copy()
    sol.columns = [str(c).strip() for c in sol.columns]

    if schedule_key_col not in sched.columns:
        raise ValueError(f"Schedule missing join key '{schedule_key_col}'")
    if solution_key_col not in sol.columns:
        raise ValueError(f"Solutions missing join key '{solution_key_col}'")

    valid_keys = set(sched[schedule_key_col].dropna().astype(str).str.strip()) - {""}
    sol_keys = sol[solution_key_col].fillna("").astype(str).str.strip()
    return sol.loc[sol_keys.isin(valid_keys)].reset_index(drop=True)


def build_schedule_control_formulas(n_rows: int) -> pd.DataFrame:
    """Generate vectorized regional COUNTIF control formula strings for rows 2..N+1."""
    if n_rows <= 0:
        return pd.DataFrame(columns=list(SCHEDULE_CONTROL_HEADERS))
    rows = pd.Series(range(2, n_rows + 2), dtype=str)
    return pd.DataFrame(
        {
            "CTRL_CENTRAL": "=COUNTIF('CTRL_CENTRAL_BASE'!AF:AF,D" + rows + ")",
            "REGULATORY_CENTRAL": "=COUNTIF('REG_CENTRAL_AUDIT'!C:C,D" + rows + ")",
            "PRELIM_CENTRAL": "=COUNTIF('PRELIM_CENTRAL'!Q:Q,D" + rows + ")",
            "CTRL_SOUTH": "=COUNTIF('CTRL_SOUTH_BASE'!AF:AF,D" + rows + ")",
            "REGULATORY_SOUTH": "=COUNTIF('REG_SOUTH_AUDIT'!C:C,D" + rows + ")",
            "PRELIM_SOUTH": "=COUNTIF('PRELIM_SOUTH'!Q:Q,D" + rows + ")",
        }
    )


def build_solution_relational_formulas(n_rows: int) -> pd.DataFrame:
    """Generate dynamic XLOOKUP/COUNTIF formula columns A-E for rows 2..N+1."""
    if n_rows <= 0:
        return pd.DataFrame(columns=list(SOLUTION_FORMULA_HEADERS))
    rows = pd.Series(range(2, n_rows + 2), dtype=str)
    return pd.DataFrame(
        {
            "INFRA_TYPE": "=XLOOKUP(LEFT(F" + rows + ",7),'STA_CENTRAL'!B:B,'STA_CENTRAL'!AB:AB)",
            "CLASSIFICATION": "=XLOOKUP(LEFT(F" + rows + ",7),'STA_CENTRAL'!B:B,'STA_CENTRAL'!CO:CO)",
            "ALT_SOLUTION": "=XLOOKUP(F" + rows + ",'PN_REG_SYNC'!E:E,'PN_REG_SYNC'!AO:AO)",
            "CROSS_CHECK": "=COUNTIF('PN_REG_SYNC'!E:E,F" + rows + ")",
            "FULL_SCHEDULE_REF": "=BZ" + rows,
        }
    )


def _write_dataframe_block(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    df: pd.DataFrame,
    start_row: int = 1,
    start_col: int = 1,
    include_header: bool = True,
) -> None:
    """Write a DataFrame block to an OpenPyXL worksheet cleanly."""
    current_row = start_row
    if include_header:
        for col_offset, col_name in enumerate(df.columns):
            ws.cell(row=current_row, column=start_col + col_offset, value=col_name)
        current_row += 1

    for row_vals in df.itertuples(index=False, name=None):
        for col_offset, val in enumerate(row_vals):
            ws.cell(row=current_row, column=start_col + col_offset, value=val)
        current_row += 1


def synchronize_workbooks(
    source_workbook: str | Path,
    output_workbook: str | Path,
    target_regions: Sequence[str] = DEFAULT_TARGET_REGIONS,
    schedule_sheet: str = "PN_SCHEDULE",
    solutions_sheet: str = "SOLUTIONS",
    dest_schedule_sheet: str = "PN_REG_SYNC",
    dest_solutions_sheet: str = "SOL_REG_SYNC",
    schedule_key_col: str = "INTERVENTION_ID",
    solution_key_col: str = "INTERVENTION_ID",
) -> dict[str, object]:
    """Run regional filtering, in-memory hash join, and dynamic formula injection."""
    wb_src = openpyxl.load_workbook(source_workbook, data_only=True, read_only=True)
    if schedule_sheet not in wb_src.sheetnames:
        wb_src.close()
        raise ValueError(f"Source sheet '{schedule_sheet}' not found in {source_workbook}")

    schedule_df = pd.read_excel(source_workbook, sheet_name=schedule_sheet)
    has_solutions = solutions_sheet in wb_src.sheetnames
    wb_src.close()

    filtered_pn = filter_schedule_by_region(schedule_df, target_regions=target_regions)
    pn_formulas = build_schedule_control_formulas(len(filtered_pn))
    pn_combined = pd.concat([filtered_pn, pn_formulas], axis=1)

    if has_solutions:
        solutions_df = pd.read_excel(source_workbook, sheet_name=solutions_sheet)
        filtered_sol = hash_join_solutions(
            filtered_pn,
            solutions_df,
            schedule_key_col=schedule_key_col,
            solution_key_col=solution_key_col,
        )
    else:
        filtered_sol = pd.DataFrame()

    sol_formulas = build_solution_relational_formulas(len(filtered_sol))

    out_path = Path(output_workbook)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb_out = openpyxl.Workbook()
    ws_pn = wb_out.active
    ws_pn.title = dest_schedule_sheet
    _write_dataframe_block(ws_pn, pn_combined, start_row=1, start_col=1, include_header=True)

    ws_sol = wb_out.create_sheet(title=dest_solutions_sheet)
    # Write relational formula columns A..E (columns 1..5) and raw solutions payload starting at F (column 6)
    _write_dataframe_block(ws_sol, sol_formulas, start_row=1, start_col=1, include_header=True)
    if not filtered_sol.empty:
        _write_dataframe_block(ws_sol, filtered_sol, start_row=1, start_col=6, include_header=True)

    wb_out.save(out_path)
    wb_out.close()

    return {
        "output_path": str(out_path),
        "pn_rows": len(filtered_pn),
        "pn_data_cols": len(filtered_pn.columns),
        "sol_rows": len(filtered_sol),
        "has_solutions": has_solutions,
    }


def run_xlwings_macro(
    source_workbook: str | Path,
    target_regions: Sequence[str] = DEFAULT_TARGET_REGIONS,
) -> dict[str, object]:
    """Optional Excel-attached execution via xlwings Book.caller()."""
    import ctypes
    import xlwings as xw

    wb = xw.Book.caller()
    schedule_df = pd.read_excel(source_workbook, sheet_name="PN_SCHEDULE")
    solutions_df = pd.read_excel(source_workbook, sheet_name="SOLUTIONS")

    filtered_pn = filter_schedule_by_region(schedule_df, target_regions=target_regions)
    pn_combined = pd.concat([filtered_pn, build_schedule_control_formulas(len(filtered_pn))], axis=1)

    filtered_sol = hash_join_solutions(filtered_pn, solutions_df)
    sol_formulas = build_solution_relational_formulas(len(filtered_sol))

    ws_pn = wb.sheets["PN_REG_SYNC"]
    ws_pn.clear()
    ws_pn.range("A1").options(index=False).value = pn_combined

    ws_sol = wb.sheets["SOL_REG_SYNC"]
    ws_sol.clear()
    ws_sol.range("A1").options(index=False).value = sol_formulas
    if not filtered_sol.empty:
        ws_sol.range("F1").options(index=False).value = filtered_sol

    ctypes.windll.user32.MessageBoxW(
        0,
        f"Sync complete: {len(filtered_pn)} schedule rows and {len(filtered_sol)} solution rows.",
        "Dynamic Formula Synchronization",
        0,
    )
    return {"pn_rows": len(filtered_pn), "sol_rows": len(filtered_sol)}


def main() -> None:
    parser = argparse.ArgumentParser(description="In-Memory Hash-Join & Dynamic Excel Formula Injection")
    parser.add_argument("--source", required=True, help="Secondary rollout workbook path (.xlsx)")
    parser.add_argument(
        "--output",
        default="data/output/synchronized_rollout_control.xlsx",
        help="Output synchronized workbook path",
    )
    parser.add_argument(
        "--regions",
        nargs="+",
        default=list(DEFAULT_TARGET_REGIONS),
        help="Target region filters",
    )
    args = parser.parse_args()

    result = synchronize_workbooks(
        source_workbook=args.source,
        output_workbook=args.output,
        target_regions=args.regions,
    )
    print(
        f"Synchronized -> {result['output_path']} "
        f"(PN rows: {result['pn_rows']}, Solution rows: {result['sol_rows']})"
    )


if __name__ == "__main__":
    main()
