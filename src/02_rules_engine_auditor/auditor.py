"""Database Auditing & Business Rules Engine.

Vectorized regulatory compliance auditor that disambiguates duplicate Excel
column headers by positional index, evaluates stage readiness across New and
Reuse deployment categories, enforces atomic site-level approval across
multi-sector records via groupby().transform('all'), and flags both pending
promotions and certified regressions.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any, Sequence

import numpy as np
import openpyxl
import pandas as pd

DEFAULT_CATEGORY_MAP: dict[str, str] = {
    "Remote_Sector_Reuse": "Reuse",
    "Remote_Sector_SN": "Reuse",
    "Reuse": "Reuse",
    "New": "New",
}

REQUIRED_FILTER_COLUMNS: tuple[str, ...] = (
    "SI_STATUS",
    "ROLLOUT_STATUS",
    "AUDIT_OWNER_STATUS",
    "STATION_STATE",
    "FLOW_CATEGORY",
    "IW_STAGE",
    "SITE_ID",
)

OUTPUT_COLUMNS: tuple[str, ...] = (
    "SITE_ID",
    "RECORD_KEY",
    "SOLUTION",
    "OBSERVATION",
    "IW_STAGE",
    "FLOW_CATEGORY",
    "AUDIT_OWNER_STATUS",
    "SUGGESTED_ACTION",
)


def resolve_duplicate_headers(
    raw_headers: Sequence[Any],
    primary_col_name: str = "OBSERVATION",
    primary_index: int = 5,
) -> list[str]:
    """Disambiguate duplicate header columns by retaining only the designated index.

    In legacy operational exports, two columns share the same header name
    (e.g., OBSERVATION at index 5 / Column F and a secondary trailing notes column).
    This preserves the canonical primary_index column and renames duplicates so
    Pandas DataFrame indexing stays unambiguous.
    """
    resolved: list[str] = []
    seen_other: dict[str, int] = {}

    for idx, col_name in enumerate(raw_headers):
        clean_name = str(col_name).strip() if col_name is not None and str(col_name).strip() != "" else f"COL_{idx}"
        if clean_name == primary_col_name:
            if idx == primary_index:
                resolved.append(primary_col_name)
            else:
                resolved.append(f"{primary_col_name}_DISCARD_{idx}")
        else:
            count = seen_other.get(clean_name, 0)
            resolved.append(clean_name if count == 0 else f"{clean_name}_DUP_{idx}")
            seen_other[clean_name] = count + 1
    return resolved


def extract_stage_number(stage_series: pd.Series) -> pd.Series:
    """Vectorially extract leading integer stage prefix (e.g., '21 - Release' -> 21)."""
    extracted = stage_series.fillna("").astype(str).str.extract(r"^\s*(\d+)", expand=False)
    return pd.to_numeric(extracted, errors="coerce").fillna(-1).astype(int)


def audit_compliance(
    df: pd.DataFrame,
    new_stage_bounds: tuple[int, int] = (18, 22),
    reuse_stage_bounds: tuple[int, int] = (7, 16),
) -> tuple[pd.DataFrame, dict[str, int]]:
    """Run vectorized compliance rules and atomic site-level approval."""
    work = df.copy()
    work.columns = [str(c).strip() for c in work.columns]

    required = set(REQUIRED_FILTER_COLUMNS) | (set(OUTPUT_COLUMNS) - {"SUGGESTED_ACTION"})
    missing = sorted(col for col in required if col not in work.columns)
    if missing:
        raise ValueError(f"Missing required audit columns: {missing}")

    for col in REQUIRED_FILTER_COLUMNS:
        work[col] = work[col].fillna("").astype(str).str.strip()

    # Step 1: Eligibility baseline filter
    eligible_mask = (
        (work["SI_STATUS"] == "OK")
        & (work["ROLLOUT_STATUS"] == "OK")
        & (work["AUDIT_OWNER_STATUS"].isin(["PENDING", "CERTIFIED"]))
        & (work["STATION_STATE"].isin(["Mobilized", "In_Formalization"]))
    )
    filtered = work.loc[eligible_mask].copy()
    if filtered.empty:
        return pd.DataFrame(columns=list(OUTPUT_COLUMNS)), {
            "eligible_rows": 0,
            "promote_count": 0,
            "regression_count": 0,
        }

    # Step 2: Vectorized flow normalization and stage-range compliance (no row loops)
    norm_category = filtered["FLOW_CATEGORY"].map(DEFAULT_CATEGORY_MAP).fillna("Other")
    stage_num = extract_stage_number(filtered["IW_STAGE"])

    is_new_valid = (norm_category == "New") & stage_num.between(new_stage_bounds[0], new_stage_bounds[1])
    is_reuse_valid = (norm_category == "Reuse") & stage_num.between(reuse_stage_bounds[0], reuse_stage_bounds[1])
    filtered["ROW_COMPLIANT"] = is_new_valid | is_reuse_valid

    # Step 3: Atomic Site-Level Rollup (100% of sectors for a SITE_ID must pass)
    filtered["SITE_APPROVED"] = filtered.groupby("SITE_ID")["ROW_COMPLIANT"].transform("all")

    # Step 4: Dual Audit Classification (Promotions vs Regressions)
    promote_mask = (filtered["AUDIT_OWNER_STATUS"] == "PENDING") & filtered["SITE_APPROVED"]
    regress_mask = (filtered["AUDIT_OWNER_STATUS"] == "CERTIFIED") & (~filtered["SITE_APPROVED"])

    audit_df = filtered.loc[promote_mask | regress_mask].copy()
    audit_df["SUGGESTED_ACTION"] = np.where(
        audit_df["AUDIT_OWNER_STATUS"] == "PENDING",
        "PROMOTE_TO_CERTIFIED",
        "AUDIT_REGRESSION_ALERT",
    )

    audit_df = audit_df.sort_values(by=["SUGGESTED_ACTION", "SITE_ID"], ascending=[False, True])
    result_df = audit_df[list(OUTPUT_COLUMNS)].reset_index(drop=True)

    summary = {
        "eligible_rows": int(len(filtered)),
        "promote_count": int((result_df["AUDIT_OWNER_STATUS"] == "PENDING").sum()),
        "regression_count": int((result_df["AUDIT_OWNER_STATUS"] == "CERTIFIED").sum()),
    }
    return result_df, summary


def load_raw_audit_sheet(
    workbook_path: str | Path,
    sheet_name: str = "COMPLIANCE_BASE",
    header_row_idx: int = 1,
    primary_observation_index: int = 5,
) -> pd.DataFrame:
    """Load an Excel sheet with multi-row offset headers and resolve duplicates."""
    wb = openpyxl.load_workbook(workbook_path, data_only=True, read_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        raise ValueError(f"Worksheet '{sheet_name}' not found in {workbook_path}")

    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if len(rows) <= header_row_idx + 1:
        raise ValueError(f"Worksheet '{sheet_name}' does not contain header and body rows.")

    raw_headers = rows[header_row_idx]
    body_rows = rows[header_row_idx + 1 :]
    resolved_headers = resolve_duplicate_headers(
        raw_headers,
        primary_col_name="OBSERVATION",
        primary_index=primary_observation_index,
    )
    return pd.DataFrame(body_rows, columns=resolved_headers)


def run_audit_file(
    input_path: str | Path,
    output_path: str | Path | None = None,
    sheet_name: str = "COMPLIANCE_BASE",
    output_sheet: str = "FILTERED_AUDIT_RESULTS",
    header_row_idx: int = 1,
    primary_observation_index: int = 5,
) -> tuple[pd.DataFrame, dict[str, int]]:
    """Run compliance audit headlessly on an .xlsx file and persist filtered results."""
    df_raw = load_raw_audit_sheet(
        workbook_path=input_path,
        sheet_name=sheet_name,
        header_row_idx=header_row_idx,
        primary_observation_index=primary_observation_index,
    )
    result_df, summary = audit_compliance(df_raw)

    target = Path(output_path) if output_path is not None else Path(input_path)
    target.parent.mkdir(parents=True, exist_ok=True)

    if target.exists() and target.resolve() == Path(input_path).resolve():
        wb = openpyxl.load_workbook(target)
        if output_sheet in wb.sheetnames:
            del wb[output_sheet]
        ws_out = wb.create_sheet(title=output_sheet)
    else:
        wb = openpyxl.Workbook()
        ws_out = wb.active
        ws_out.title = output_sheet

    ws_out.append(list(result_df.columns))
    for row in result_df.itertuples(index=False, name=None):
        ws_out.append(list(row))

    wb.save(target)
    return result_df, summary


def run_xlwings_macro(
    source_sheet: str = "COMPLIANCE_BASE",
    target_sheet: str = "FILTERED_AUDIT_RESULTS",
    header_row_idx: int = 1,
    primary_observation_index: int = 5,
) -> dict[str, int]:
    """Optional Excel-attached execution via xlwings Book.caller()."""
    import ctypes
    import xlwings as xw

    wb = xw.Book.caller()
    if source_sheet not in [s.name for s in wb.sheets]:
        ctypes.windll.user32.MessageBoxW(0, f"Sheet '{source_sheet}' not found!", "Audit Error", 0)
        return {"eligible_rows": 0, "promote_count": 0, "regression_count": 0}

    raw_values = wb.sheets[source_sheet].used_range.value
    if not raw_values or len(raw_values) <= header_row_idx + 1:
        ctypes.windll.user32.MessageBoxW(0, "Worksheet contains no operational data.", "Audit Error", 0)
        return {"eligible_rows": 0, "promote_count": 0, "regression_count": 0}

    headers = resolve_duplicate_headers(
        raw_values[header_row_idx],
        primary_col_name="OBSERVATION",
        primary_index=primary_observation_index,
    )
    df_raw = pd.DataFrame(raw_values[header_row_idx + 1 :], columns=headers)
    result_df, summary = audit_compliance(df_raw)

    if target_sheet in [s.name for s in wb.sheets]:
        dest = wb.sheets[target_sheet]
        dest.clear()
    else:
        dest = wb.sheets.add(target_sheet)

    if not result_df.empty:
        dest.range("A1").options(index=False).value = result_df
        dest.autofit()

    ctypes.windll.user32.MessageBoxW(
        0,
        (
            f"Audit completed.\n\n"
            f"-> {summary['promote_count']} rows ready for certification.\n"
            f"-> {summary['regression_count']} certified rows flagged for regression review."
        ),
        "Compliance Audit Report",
        0,
    )
    return summary


def main() -> None:
    parser = argparse.ArgumentParser(description="Vectorized Regulatory Compliance Rules Engine")
    parser.add_argument("--input", required=True, help="Source Excel workbook path")
    parser.add_argument("--output", default=None, help="Output workbook path (defaults to in-place sheet update)")
    parser.add_argument("--sheet", default="COMPLIANCE_BASE", help="Source worksheet name")
    parser.add_argument("--header-row", type=int, default=1, help="0-based header row index (default: 1 for Excel row 2)")
    args = parser.parse_args()

    _, summary = run_audit_file(
        input_path=args.input,
        output_path=args.output,
        sheet_name=args.sheet,
        header_row_idx=args.header_row,
    )
    print(
        f"Audit completed: {summary['eligible_rows']} eligible rows | "
        f"{summary['promote_count']} promotions | {summary['regression_count']} regressions"
    )


if __name__ == "__main__":
    main()
