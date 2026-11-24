"""Vectorized Logic Validator (Legacy VBA / Row-Loop Refactoring).

Performs 100% vectorized multi-tier frequency and bandwidth cross-validation
using Pandas Explode-Merge-GroupBy and NumPy boolean masking—eliminating
legacy row-by-row loops and Excel COM bottlenecks.
"""

from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd

REQUIRED_VALIDATOR_COLUMNS: tuple[str, ...] = (
    "TECH_TIER",
    "SOLUTION",
    "TIER_1",
    "TIER_2",
    "TIER_3",
    "TIER_4",
    "BW_DUAL_ACTIVITY",
    "BW_EXT_ACTIVITY",
)

_TIER_3_TOKEN_ALIASES: dict[str, str] = {
    "LOWA": "LOW",
    "HIGHA": "HIGH",
    "HIGHB": "HIGH",
    "MIDC": "MID",
}


def _clean_text_series(series: pd.Series) -> pd.Series:
    """Normalize string values vectorially and strip residual '.0' float suffixes."""
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace(r"\.0$", "", regex=True)
    )


def _normalize_tier3_tokens(tokens: pd.Series, is_tier3: pd.Series) -> pd.Series:
    """Map Tier-3 frequency sub-band aliases to canonical band identifiers."""
    mapped = tokens.replace(_TIER_3_TOKEN_ALIASES)
    return pd.Series(np.where(is_tier3, mapped, tokens), index=tokens.index)


def validate_frequency_matrix(df: pd.DataFrame) -> pd.Series:
    """Evaluate multi-tier frequency compliance completely vectorially.

    Priority hierarchy:
      1. 'Check Schedule' if SOLUTION is '#N/A' or '#N/D'.
      2. 'OK SWAP' if SOLUTION is null, empty, '0', 'NAN', or 'NONE'.
      3. 'OK' if every requested band token (slash-delimited) exists in the
         installed tier column (underscore-delimited), dual-carrier rules on
         BW_DUAL_ACTIVITY are satisfied, and TIER_3 EXT_BW extension constraints
         are met; otherwise 'NOK'.
    """
    orig_index = df.index
    work = df.reset_index(drop=True)
    work.columns = [str(col).strip() for col in work.columns]

    missing = [col for col in REQUIRED_VALIDATOR_COLUMNS if col not in work.columns]
    if missing:
        raise ValueError(f"Missing required validator columns: {missing}")

    if work.empty:
        return pd.Series([], index=orig_index, dtype=str, name="VALIDATION_STATUS")

    n_rows = len(work)
    row_pos = np.arange(n_rows)

    raw_solution = work["SOLUTION"]
    sol_clean = _clean_text_series(raw_solution)
    tech = _clean_text_series(work["TECH_TIER"])
    t1 = _clean_text_series(work["TIER_1"])
    t2 = _clean_text_series(work["TIER_2"])
    t3 = _clean_text_series(work["TIER_3"])
    t4 = _clean_text_series(work["TIER_4"])
    bw_dual = _clean_text_series(work["BW_DUAL_ACTIVITY"])
    bw_ext = _clean_text_series(work["BW_EXT_ACTIVITY"])

    # Priority 1 & 2 masks
    cond_check_schedule = sol_clean.isin(["#N/A", "#N/D"])
    cond_ok_swap = raw_solution.isna() | sol_clean.isin(["", "0", "NAN", "NONE"])

    # Select installed frequency string for each row's technology tier
    installed_raw = pd.Series(
        np.select(
            [tech == "TIER_1", tech == "TIER_2", tech == "TIER_3", tech == "TIER_4"],
            [t1, t2, t3, t4],
            default="",
        ),
        index=row_pos,
    )
    tech_eligible = (
        ~(cond_check_schedule | cond_ok_swap)
        & tech.isin(["TIER_1", "TIER_2", "TIER_3", "TIER_4"])
        & (installed_raw != "")
        & (installed_raw != "NONE")
    )

    # Row-level dual-carrier matrix (ALPHA / BETA / flexible) for TIER_2 & TIER_3
    padded_installed = "_" + installed_raw + "_"
    has_duala = padded_installed.str.contains("_DUALA_", regex=False).to_numpy()
    has_dualb = padded_installed.str.contains("_DUALB_", regex=False).to_numpy()
    has_dual_any = has_duala | has_dualb | padded_installed.str.contains("_DUAL_", regex=False).to_numpy()

    is_alpha = bw_dual.str.contains("ALPHA", regex=False).to_numpy()
    is_beta = bw_dual.str.contains("BETA", regex=False).to_numpy()

    dual_ok_by_row = np.select(
        [is_alpha & is_beta, is_alpha, is_beta],
        [has_duala & has_dualb, has_duala, has_dualb],
        default=has_dual_any,
    )

    # Row-level mandatory extension tag when TIER_3 requests EXT_BW
    ext_ok_by_row = np.where(
        (tech.to_numpy() == "TIER_3") & (bw_ext.to_numpy() == "EXT_BW"),
        padded_installed.str.contains("_MIDEXT_", regex=False).to_numpy(),
        True,
    )

    # Explode installed bands -> long DataFrame (row_id, tok)
    inst_series = pd.Series(installed_raw.to_numpy(), index=row_pos)
    inst_exploded = inst_series[tech_eligible.to_numpy()].str.split("_").explode().str.strip()
    inst_exploded = inst_exploded[inst_exploded != ""]
    if not inst_exploded.empty:
        inst_is_t3 = tech.to_numpy()[inst_exploded.index] == "TIER_3"
        inst_norm = _normalize_tier3_tokens(inst_exploded, pd.Series(inst_is_t3, index=inst_exploded.index))
        df_inst = pd.DataFrame({"row_id": inst_exploded.index.to_numpy(), "tok": inst_norm.to_numpy()}).drop_duplicates()
    else:
        df_inst = pd.DataFrame({"row_id": np.array([], dtype=int), "tok": pd.Series([], dtype=str)})

    # Explode requested solution tokens -> long DataFrame (row_id, tok)
    sol_series = pd.Series(sol_clean.to_numpy(), index=row_pos)
    req_exploded = sol_series[tech_eligible.to_numpy()].str.split("/").explode().str.strip()
    req_exploded = req_exploded[req_exploded != ""]

    if not req_exploded.empty:
        req_row_ids = req_exploded.index.to_numpy()
        req_is_t3 = tech.to_numpy()[req_row_ids] == "TIER_3"
        req_is_dual_tier = np.isin(tech.to_numpy()[req_row_ids], ["TIER_2", "TIER_3"])
        req_norm = _normalize_tier3_tokens(req_exploded, pd.Series(req_is_t3, index=req_exploded.index))

        df_req = pd.DataFrame({"row_id": req_row_ids, "tok": req_norm.to_numpy()})
        merged = pd.merge(df_req, df_inst, on=["row_id", "tok"], how="left", indicator=True, validate="many_to_one")

        is_dual_token = req_is_dual_tier & df_req["tok"].str.contains("DUAL", regex=False).to_numpy()
        exact_match = (merged["_merge"] == "both").to_numpy()
        dual_match = dual_ok_by_row[req_row_ids]
        token_ok = np.where(is_dual_token, dual_match, exact_match)

        all_tokens_ok = (
            pd.Series(token_ok, index=req_row_ids)
            .groupby(level=0)
            .all()
            .reindex(row_pos, fill_value=False)
            .to_numpy()
        )
    else:
        all_tokens_ok = np.zeros(n_rows, dtype=bool)

    cond_tech_ok = tech_eligible.to_numpy() & all_tokens_ok & ext_ok_by_row

    result_array = np.select(
        [cond_check_schedule.to_numpy(), cond_ok_swap.to_numpy(), cond_tech_ok],
        ["Check Schedule", "OK SWAP", "OK"],
        default="NOK",
    )
    return pd.Series(result_array, index=orig_index, name="VALIDATION_STATUS")


def run_validator_file(
    input_path: str | Path,
    output_path: str | Path | None = None,
    sheet_name: str | None = None,
    header_row_idx: int = 1,
    target_column_letter: str = "T",
) -> pd.DataFrame:
    """Validate an Excel workbook headlessly and inject statuses into target column."""
    wb = openpyxl.load_workbook(input_path)
    ws = wb[sheet_name] if sheet_name else wb.active

    rows = list(ws.iter_rows(values_only=True))
    if len(rows) <= header_row_idx + 1:
        wb.close()
        raise ValueError("Worksheet has no data rows beneath the header.")

    headers = [str(h).strip() if h is not None else f"COL_{i}" for i, h in enumerate(rows[header_row_idx])]
    df = pd.DataFrame(rows[header_row_idx + 1 :], columns=headers)
    statuses = validate_frequency_matrix(df)
    df["VALIDATION_STATUS"] = statuses

    target_col_idx = openpyxl.utils.column_index_from_string(target_column_letter)
    start_excel_row = header_row_idx + 2
    for offset, val in enumerate(statuses.tolist()):
        ws.cell(row=start_excel_row + offset, column=target_col_idx, value=val)

    dest = Path(output_path) if output_path is not None else Path(input_path)
    dest.parent.mkdir(parents=True, exist_ok=True)
    wb.save(dest)
    wb.close()
    return df


def run_xlwings_macro(header_row: int = 2, target_column: str = "T") -> int:
    """Optional Excel-attached execution via xlwings Book.caller()."""
    import ctypes
    import xlwings as xw

    wb = xw.Book.caller()
    sheet = wb.sheets.active
    last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end("up").row
    if last_row <= header_row:
        return 0

    df = sheet.range(f"A{header_row}:AG{last_row}").options(pd.DataFrame, index=False, header=True).value
    statuses = validate_frequency_matrix(df)
    sheet.range(f"{target_column}{header_row + 1}").options(transpose=True).value = statuses.tolist()
    ctypes.windll.user32.MessageBoxW(0, f"Validated {len(statuses)} rows.", "Vectorized Logic Validator", 0)
    return len(statuses)


def main() -> None:
    parser = argparse.ArgumentParser(description="Vectorized Multi-Tier Frequency Matrix Validator")
    parser.add_argument("--input", required=True, help="Input Excel workbook path")
    parser.add_argument("--output", default=None, help="Output workbook path (defaults to in-place update)")
    parser.add_argument("--sheet", default=None, help="Target worksheet name (defaults to active sheet)")
    parser.add_argument("--header-row", type=int, default=1, help="0-based header row index (default: 1)")
    args = parser.parse_args()

    df_out = run_validator_file(
        input_path=args.input,
        output_path=args.output,
        sheet_name=args.sheet,
        header_row_idx=args.header_row,
    )
    counts = df_out["VALIDATION_STATUS"].value_counts().to_dict()
    print(f"Validation completed ({len(df_out)} rows): {counts}")


if __name__ == "__main__":
    main()
