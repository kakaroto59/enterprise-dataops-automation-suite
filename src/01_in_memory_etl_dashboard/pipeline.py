"""In-Memory ETL Pipeline & CLI SLA Dashboard.

Ingests multi-day semicolon-delimited raw deployment logs (New and Reuse flows),
merges them in RAM with a master rollout schedule, computes vectorized stage SLA
dwell metrics (excluding terminal administrative stages from active cycle sums),
and exports a formatted Excel workbook with native bar charts.
"""

from __future__ import annotations

import argparse
import re
from datetime import date, datetime
from pathlib import Path
from typing import Iterable, Sequence

import numpy as np
import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

NEW_FLOW_STAGES: tuple[str, ...] = (
    "21 - Logical Acceptance",
    "22 - Traffic Release",
    "23 - Final Site Acceptance",
)

REUSE_FLOW_STAGES: tuple[str, ...] = (
    "22 - Execute Field Test",
    "23 - Approve Field Test",
    "24 - Submit Acceptance CRQ",
    "25 - Conditional Network Acceptance",
    "25.1 - Resolve Blocking Issues",
    "26 - Traffic Release",
    "27 - Final Element Acceptance",
)

TERMINAL_STAGES: frozenset[str] = frozenset(
    {
        "23 - Final Site Acceptance",
        "27 - Final Element Acceptance",
    }
)

_FILENAME_DATE_PATTERN = re.compile(r"(\d{8})-[^\\/]+$")


def extract_date_from_filename(filepath: str | Path) -> date | None:
    """Extract snapshot date (ddmmyyyy) immediately preceding the final hyphen."""
    name = Path(filepath).name
    match = _FILENAME_DATE_PATTERN.search(name)
    if not match:
        return None
    try:
        return datetime.strptime(match.group(1), "%d%m%Y").date()
    except ValueError:
        return None


def _normalize_single_log(
    file_path: str | Path,
    flow_type: str,
    sep: str = ";",
) -> pd.DataFrame:
    """Read one raw log file into memory and tag snapshot metadata vectorially."""
    path = Path(file_path)
    df = pd.read_csv(path, sep=sep, encoding="utf-8", dtype=str)
    df.columns = [str(col).strip().upper() for col in df.columns]

    snapshot_dt = extract_date_from_filename(path) or date.today()
    df["SNAPSHOT_DATE"] = pd.Timestamp(snapshot_dt)
    df["FLOW_TYPE"] = flow_type

    station_col = df["STATION"] if "STATION" in df.columns else pd.Series("", index=df.index)
    activity_col = df["ACTIVITY"] if "ACTIVITY" in df.columns else pd.Series("", index=df.index)

    station_clean = station_col.fillna("").astype(str).str.strip()
    activity_clean = activity_col.fillna("").astype(str).str.strip()

    # Reuse logs key stations by ACTIVITY; New logs prioritize STATION with ACTIVITY fallback.
    resolved_site = np.where(
        flow_type == "Reuse",
        activity_clean,
        np.where(station_clean != "", station_clean, activity_clean),
    )
    df["SITE_ID"] = pd.Series(resolved_site, index=df.index).astype(str).str.strip()

    if "STAGE_NAME" not in df.columns:
        df["STAGE_NAME"] = ""
    df["STAGE_NAME"] = df["STAGE_NAME"].fillna("").astype(str).str.strip()

    if "DAYS_IN_STAGE" in df.columns:
        df["DAYS_IN_STAGE"] = pd.to_numeric(df["DAYS_IN_STAGE"], errors="coerce").fillna(0.0)
    else:
        df["DAYS_IN_STAGE"] = 0.0

    return df


def consolidate_raw_logs(
    log_entries: Sequence[tuple[str | Path, str]],
    sep: str = ";",
) -> pd.DataFrame:
    """Concatenate multiple daily New/Reuse raw logs completely in memory."""
    if not log_entries:
        return pd.DataFrame(
            columns=["SNAPSHOT_DATE", "FLOW_TYPE", "SITE_ID", "STAGE_NAME", "DAYS_IN_STAGE"]
        )

    frames = [_normalize_single_log(path, flow_type, sep=sep) for path, flow_type in log_entries]
    consolidated = pd.concat(frames, ignore_index=True)

    # Deduplicate overlapping re-runs by keeping the latest snapshot per site/flow/stage.
    consolidated = consolidated.sort_values("SNAPSHOT_DATE").drop_duplicates(
        subset=["SNAPSHOT_DATE", "FLOW_TYPE", "SITE_ID", "STAGE_NAME"],
        keep="last",
    )
    return consolidated.reset_index(drop=True)


def filter_and_merge_schedule(
    raw_df: pd.DataFrame,
    schedule_df: pd.DataFrame,
    target_sites: Iterable[str] | None = None,
    target_regions: Sequence[str] = ("REG_CENTRAL", "REG_SOUTH"),
) -> pd.DataFrame:
    """Filter master rollout schedule by region/sites and merge in-memory with raw logs."""
    if raw_df.empty:
        return raw_df.copy()

    sched = schedule_df.copy()
    sched.columns = [str(col).strip().upper() for col in sched.columns]

    if "SITE_ID" not in sched.columns and "STATION" in sched.columns:
        sched["SITE_ID"] = sched["STATION"]

    sched["SITE_ID"] = sched["SITE_ID"].fillna("").astype(str).str.strip()

    if "REGIONAL" in sched.columns and target_regions:
        sched["REGIONAL"] = sched["REGIONAL"].fillna("").astype(str).str.strip().str.upper()
        region_set = {r.strip().upper() for r in target_regions}
        sched = sched[sched["REGIONAL"].isin(region_set)]

    if target_sites is not None:
        site_set = {str(s).strip() for s in target_sites if str(s).strip()}
        if site_set:
            raw_df = raw_df[raw_df["SITE_ID"].isin(site_set)]
            sched = sched[sched["SITE_ID"].isin(site_set)]

    sched_dedup = sched.drop_duplicates(subset=["SITE_ID"], keep="first")
    keep_cols = [c for c in ["SITE_ID", "REGIONAL", "PRIORITY_TIER", "VENDOR_TEAM"] if c in sched_dedup.columns]

    merged = pd.merge(
        raw_df,
        sched_dedup[keep_cols],
        on="SITE_ID",
        how="inner" if not sched_dedup.empty else "left",
    )
    return merged.reset_index(drop=True)


def compute_stage_slas(
    merged_df: pd.DataFrame,
    terminal_stages: frozenset[str] = TERMINAL_STAGES,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Compute vectorized per-stage mean dwell times and active cycle totals."""
    if merged_df.empty:
        empty_summary = pd.DataFrame(
            columns=["FLOW_TYPE", "STAGE_NAME", "AVG_DAYS", "MAX_DAYS", "SITE_OBSERVATIONS", "INCLUDED_IN_CYCLE"]
        )
        empty_totals = pd.DataFrame(columns=["FLOW_TYPE", "TOTAL_ACTIVE_CYCLE_DAYS", "STAGE_COUNT"])
        return merged_df.copy(), merged_df.copy(), empty_summary, empty_totals

    df = merged_df.copy()
    df["INCLUDED_IN_CYCLE"] = ~df["STAGE_NAME"].isin(terminal_stages)

    stage_summary = (
        df.groupby(["FLOW_TYPE", "STAGE_NAME", "INCLUDED_IN_CYCLE"], as_index=False)
        .agg(
            AVG_DAYS=("DAYS_IN_STAGE", "mean"),
            MAX_DAYS=("DAYS_IN_STAGE", "max"),
            SITE_OBSERVATIONS=("SITE_ID", "nunique"),
        )
        .sort_values(["FLOW_TYPE", "STAGE_NAME"])
        .reset_index(drop=True)
    )
    stage_summary["AVG_DAYS"] = stage_summary["AVG_DAYS"].round(2)
    stage_summary["MAX_DAYS"] = stage_summary["MAX_DAYS"].round(2)

    active_stages = stage_summary[stage_summary["INCLUDED_IN_CYCLE"]]
    cycle_totals = (
        active_stages.groupby("FLOW_TYPE", as_index=False)
        .agg(
            TOTAL_ACTIVE_CYCLE_DAYS=("AVG_DAYS", "sum"),
            STAGE_COUNT=("STAGE_NAME", "count"),
        )
        .reset_index(drop=True)
    )
    cycle_totals["TOTAL_ACTIVE_CYCLE_DAYS"] = cycle_totals["TOTAL_ACTIVE_CYCLE_DAYS"].round(2)

    new_df = df[df["FLOW_TYPE"] == "New"].reset_index(drop=True)
    reuse_df = df[df["FLOW_TYPE"] == "Reuse"].reset_index(drop=True)
    return new_df, reuse_df, stage_summary, cycle_totals


def _write_df_to_sheet(ws: openpyxl.worksheet.worksheet.Worksheet, df: pd.DataFrame) -> None:
    """Populate an OpenPyXL worksheet from a DataFrame with styled header row."""
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    for row_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=1):
        ws.append(list(row))
        if row_idx == 1:
            for col_idx in range(1, len(row) + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.fill = header_fill
                cell.font = header_font


def _add_flow_bar_chart(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    title: str,
    min_row: int,
    max_row: int,
    anchor: str,
) -> None:
    """Attach a horizontal/vertical bar chart of average stage SLA days."""
    if max_row < min_row:
        return
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = title
    chart.y_axis.title = "Average Days"
    chart.x_axis.title = "Workflow Stage"

    data_ref = Reference(ws, min_col=3, min_row=min_row - 1, max_row=max_row)
    cats_ref = Reference(ws, min_col=2, min_row=min_row, max_row=max_row)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.legend = None
    chart.width = 18
    chart.height = 10
    ws.add_chart(chart, anchor)


def export_dashboard_workbook(
    new_df: pd.DataFrame,
    reuse_df: pd.DataFrame,
    stage_summary_df: pd.DataFrame,
    cycle_totals_df: pd.DataFrame,
    output_path: str | Path,
) -> Path:
    """Write consolidated details and SLA bar-chart dashboard to an Excel workbook."""
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()
    ws_dash = wb.active
    ws_dash.title = "SLA_Dashboard"

    ws_new = wb.create_sheet(title="New_Flow_Details")
    ws_reuse = wb.create_sheet(title="Reuse_Flow_Details")

    _write_df_to_sheet(ws_new, new_df)
    _write_df_to_sheet(ws_reuse, reuse_df)

    # Populate Dashboard sheet: Stage Summary table + Cycle Totals side table
    summary_cols = ["FLOW_TYPE", "STAGE_NAME", "AVG_DAYS", "MAX_DAYS", "SITE_OBSERVATIONS", "INCLUDED_IN_CYCLE"]
    summary_export = stage_summary_df[summary_cols] if not stage_summary_df.empty else pd.DataFrame(columns=summary_cols)
    _write_df_to_sheet(ws_dash, summary_export)

    totals_start_row = len(summary_export) + 4
    ws_dash.cell(row=totals_start_row, column=1, value="ACTIVE_CYCLE_SLA_SUMMARY").font = Font(bold=True)
    for offset_r, row in enumerate(
        dataframe_to_rows(cycle_totals_df, index=False, header=True),
        start=totals_start_row + 1,
    ):
        for col_idx, val in enumerate(row, start=1):
            cell = ws_dash.cell(row=offset_r, column=col_idx, value=val)
            if offset_r == totals_start_row + 1:
                cell.font = Font(bold=True)

    # Add separate bar charts for New and Reuse segments when present
    if not summary_export.empty:
        new_mask = summary_export["FLOW_TYPE"] == "New"
        reuse_mask = summary_export["FLOW_TYPE"] == "Reuse"

        if new_mask.any():
            indices = summary_export.index[new_mask]
            _add_flow_bar_chart(
                ws_dash,
                title="New Flow - Average Dwell Days per Stage",
                min_row=int(indices.min()) + 2,
                max_row=int(indices.max()) + 2,
                anchor="H2",
            )
        if reuse_mask.any():
            indices = summary_export.index[reuse_mask]
            _add_flow_bar_chart(
                ws_dash,
                title="Reuse Flow - Average Dwell Days per Stage",
                min_row=int(indices.min()) + 2,
                max_row=int(indices.max()) + 2,
                anchor="H22",
            )

    wb.save(out)
    return out


def run_etl_pipeline(
    log_entries: Sequence[tuple[str | Path, str]],
    schedule_path: str | Path,
    output_path: str | Path,
    sites_file: str | Path | None = None,
    target_regions: Sequence[str] = ("REG_CENTRAL", "REG_SOUTH"),
) -> dict[str, object]:
    """Execute end-to-end In-Memory ETL and generate the Excel SLA dashboard."""
    raw_df = consolidate_raw_logs(log_entries)
    schedule_df = pd.read_excel(schedule_path)

    target_sites: list[str] | None = None
    if sites_file is not None:
        sites_df = pd.read_csv(sites_file) if str(sites_file).lower().endswith(".csv") else pd.read_excel(sites_file)
        first_col = sites_df.columns[0]
        target_sites = sites_df[first_col].dropna().astype(str).str.strip().tolist()

    merged_df = filter_and_merge_schedule(
        raw_df=raw_df,
        schedule_df=schedule_df,
        target_sites=target_sites,
        target_regions=target_regions,
    )
    new_df, reuse_df, summary_df, totals_df = compute_stage_slas(merged_df)
    saved_path = export_dashboard_workbook(new_df, reuse_df, summary_df, totals_df, output_path)

    return {
        "output_path": str(saved_path),
        "raw_rows": len(raw_df),
        "merged_rows": len(merged_df),
        "new_rows": len(new_df),
        "reuse_rows": len(reuse_df),
        "cycle_totals": totals_df.to_dict(orient="records"),
    }


def _parse_log_arg(spec: str) -> tuple[Path, str]:
    if "=" not in spec:
        raise argparse.ArgumentTypeError(f"Expected FLOW=PATH format, got: {spec}")
    flow, path_str = spec.split("=", 1)
    flow_norm = flow.strip().capitalize()
    if flow_norm not in {"New", "Reuse"}:
        raise argparse.ArgumentTypeError(f"Flow must be New or Reuse, got: {flow}")
    return Path(path_str.strip()), flow_norm


def main() -> None:
    parser = argparse.ArgumentParser(description="In-Memory Telecom Rollout ETL & SLA Dashboard")
    parser.add_argument(
        "--log",
        dest="logs",
        action="append",
        required=True,
        type=_parse_log_arg,
        help="Input log in FLOW=PATH format (e.g., New=data/mock/23112026-new.csv)",
    )
    parser.add_argument("--schedule", required=True, help="Path to master rollout schedule (.xlsx)")
    parser.add_argument("--sites", default=None, help="Optional CSV/Excel containing target SITE_ID list")
    parser.add_argument(
        "--regions",
        nargs="+",
        default=["REG_CENTRAL", "REG_SOUTH"],
        help="Regional codes to include from master schedule",
    )
    parser.add_argument(
        "--output",
        default="data/output/consolidated_sla_dashboard.xlsx",
        help="Destination Excel dashboard path",
    )
    args = parser.parse_args()

    result = run_etl_pipeline(
        log_entries=args.logs,
        schedule_path=args.schedule,
        output_path=args.output,
        sites_file=args.sites,
        target_regions=args.regions,
    )
    print(f"ETL completed -> {result['output_path']} ({result['merged_rows']} rows processed)")


if __name__ == "__main__":
    main()
