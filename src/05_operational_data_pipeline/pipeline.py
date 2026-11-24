"""Single-Table Operational Data Pipeline & Dual-Stream Log Engine.

Evolved from a multi-list relational model (which suffered from N+1 query
bottlenecks and delegation caps) into a denormalized Single-Table architecture
with embedded Dual-Stream logs ([[EVENTS]] for system audit trails & attention
flags vs. [[INTERACTIONS]] for executive follow-ups). Produces vectorized KPIs,
Division x Priority matrices, JSON feeds, and an interactive HTML Hub preview.
"""

from __future__ import annotations

import argparse
from html import escape
import json
from pathlib import Path
import re

import numpy as np
import pandas as pd

_CLOSED_ALIASES: frozenset[str] = frozenset({"CLOSED", "COMPLETED", "DONE", "ENCERRADO", "FINALIZADO"})
_BOARD_ALIASES: frozenset[str] = frozenset({"BOARD", "EXECUTIVE_BOARD", "CD"})


def parse_dual_stream_history(history_series: pd.Series) -> pd.DataFrame:
    """Vectorially separate [[EVENTS]] and [[INTERACTIONS]] inside a single log field.

    Ensures automated system events (e.g., priority switches or attention flags)
    never overwrite the executive 'LATEST_INTERACTION' shown in summary tables.
    """
    raw = history_series.fillna("").astype(str)
    pos_index = pd.RangeIndex(len(raw))
    raw_pos = pd.Series(raw.to_numpy(), index=pos_index)

    events_block = (
        raw_pos.str.extract(r"\[\[EVENTS\]\]\s*(.*?)(?=\[\[INTERACTIONS\]\]|\Z)", flags=re.DOTALL, expand=False)
        .fillna("")
        .str.strip()
    )
    inter_block = (
        raw_pos.str.extract(r"\[\[INTERACTIONS\]\]\s*(.*?)(?=\[\[EVENTS\]\]|\Z)", flags=re.DOTALL, expand=False)
        .fillna("")
        .str.strip()
    )

    # Count non-empty lines per stream vectorially
    ev_lines = events_block.str.split(r"\r?\n").explode().str.strip()
    ev_lines = ev_lines[ev_lines != ""]
    event_counts = (
        ev_lines.groupby(level=0).size().reindex(pos_index, fill_value=0).astype(int)
        if not ev_lines.empty
        else pd.Series(0, index=pos_index, dtype=int)
    )

    int_lines = inter_block.str.split(r"\r?\n").explode().str.strip()
    int_lines = int_lines[int_lines != ""]

    if not int_lines.empty:
        inter_counts = int_lines.groupby(level=0).size().reindex(pos_index, fill_value=0).astype(int)
        parsed_dates = pd.to_datetime(int_lines.str.extract(r"(\d{4}-\d{2}-\d{2})", expand=False), errors="coerce")
        line_df = pd.DataFrame(
            {
                "text": int_lines.to_numpy(),
                "dt": parsed_dates.to_numpy(),
                "order": np.arange(len(int_lines)),
            },
            index=int_lines.index,
        )
        latest_df = (
            line_df.sort_values(["dt", "order"], na_position="first")
            .groupby(level=0)
            .last()
            .reindex(pos_index)
        )
        latest_text = latest_df["text"].fillna("").astype(str)
        latest_date = pd.to_datetime(latest_df["dt"], errors="coerce")
    else:
        inter_counts = pd.Series(0, index=pos_index, dtype=int)
        latest_text = pd.Series("", index=pos_index, dtype=str)
        latest_date = pd.Series(pd.NaT, index=pos_index)

    return pd.DataFrame(
        {
            "EVENTS_LOG": events_block.to_numpy(),
            "INTERACTIONS_LOG": inter_block.to_numpy(),
            "EVENT_COUNT": event_counts.to_numpy(),
            "INTERACTION_COUNT": inter_counts.to_numpy(),
            "LATEST_INTERACTION": latest_text.to_numpy(),
            "LAST_INTERACTION_DATE": latest_date.to_numpy(),
        },
        index=history_series.index,
    )


def sanitize_single_table_processes(raw_df: pd.DataFrame) -> pd.DataFrame:
    """Sanitize and normalize single-table operational records vectorially."""
    work = raw_df.copy()
    work.columns = [str(col).strip().upper() for col in work.columns]

    required = {"CASE_NUMBER", "THEME", "PRIORITY", "DIVISION_UNIT", "ANALYSIS_STATUS"}
    missing = sorted(required - set(work.columns))
    if missing:
        raise ValueError(f"Missing required single-table columns: {missing}")

    # Explode multi-case cells vectorially (no .apply or .iterrows)
    raw_case = work["CASE_NUMBER"].fillna("").astype(str)
    cleaned_case = (
        raw_case.str.replace("\u00a0", "", regex=False)
        .str.replace("\u200b", "", regex=False)
        .str.strip()
        .str.upper()
    )
    work["CASE_NUMBER"] = cleaned_case.str.split(r"[,;]\s*")
    work = work.explode("CASE_NUMBER")
    work["CASE_NUMBER"] = work["CASE_NUMBER"].fillna("").astype(str).str.strip()
    work = work[~work["CASE_NUMBER"].isin(["", "NAN", "NONE", "NULL"])].drop_duplicates(
        subset=["CASE_NUMBER"], keep="first"
    )
    work = work.reset_index(drop=True)

    # Clean division acronym (strip trailing full names)
    work["DIVISION_UNIT"] = (
        work["DIVISION_UNIT"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .str.extract(r"^([A-Z0-9_]+)", expand=False)
        .fillna("OTHER")
    )

    # Strict numeric priority ('1', '2', '3')
    work["PRIORITY"] = (
        work["PRIORITY"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.extract(r"^([123])", expand=False)
        .fillna("3")
    )

    # Two-state operational status ('In Progress' / 'Closed')
    status_upper = work["ANALYSIS_STATUS"].fillna("").astype(str).str.strip().str.upper()
    work["ANALYSIS_STATUS"] = np.where(status_upper.isin(_CLOSED_ALIASES), "Closed", "In Progress")

    # Board escalation flag
    cls_col = work["CLASSIFICATION"] if "CLASSIFICATION" in work.columns else pd.Series("DIVISION", index=work.index)
    work["ESCALATED_TO_BOARD"] = cls_col.fillna("").astype(str).str.strip().str.upper().isin(_BOARD_ALIASES)

    # Optional columns normalization
    for opt_col, default_val in [
        ("SYNTHESIS", ""),
        ("INTERNAL_OWNER", ""),
        ("EXTERNAL_CONTACT", ""),
        ("CASE_LINK", ""),
    ]:
        if opt_col not in work.columns:
            work[opt_col] = default_val
        else:
            work[opt_col] = work[opt_col].fillna(default_val).astype(str).str.strip()

    if "EXTERNAL_ATTENTION_FLAG" in work.columns:
        work["EXTERNAL_ATTENTION_FLAG"] = (
            work["EXTERNAL_ATTENTION_FLAG"]
            .fillna(False)
            .astype(str)
            .str.strip()
            .str.upper()
            .isin(["TRUE", "1", "YES", "SIM"])
        )
    else:
        work["EXTERNAL_ATTENTION_FLAG"] = False

    history_series = work["INTERNAL_HISTORY"] if "INTERNAL_HISTORY" in work.columns else pd.Series("", index=work.index)
    parsed_logs = parse_dual_stream_history(history_series)

    work = pd.concat([work, parsed_logs], axis=1)
    work["THEME"] = work["THEME"].fillna("Untitled Case").astype(str).str.strip()
    return work


def compute_executive_kpis_and_matrix(
    df: pd.DataFrame,
    reference_date: str | pd.Timestamp | None = None,
) -> tuple[dict[str, object], pd.DataFrame]:
    """Compute the 4 single-table executive KPIs and Division x Priority matrix."""
    ref_ts = pd.Timestamp(reference_date).normalize() if reference_date is not None else pd.Timestamp.now().normalize()

    total_monitored = int(len(df))
    in_progress_count = int((df["ANALYSIS_STATUS"] == "In Progress").sum()) if total_monitored else 0
    closed_count = int((df["ANALYSIS_STATUS"] == "Closed").sum()) if total_monitored else 0
    high_priority_p1 = int((df["PRIORITY"] == "1").sum()) if total_monitored else 0
    external_requests = int(df["EXTERNAL_ATTENTION_FLAG"].sum()) if total_monitored else 0

    if total_monitored:
        is_active = df["ANALYSIS_STATUS"] != "Closed"
        last_dt = pd.to_datetime(df["LAST_INTERACTION_DATE"], errors="coerce")
        days_since = (ref_ts - last_dt).dt.days
        stale_mask = is_active & (days_since.isna() | (days_since >= 10))
        stale_10d_count = int(stale_mask.sum())
    else:
        stale_10d_count = 0

    kpis: dict[str, object] = {
        "total_monitored": total_monitored,
        "in_progress_count": in_progress_count,
        "closed_count": closed_count,
        "high_priority_p1_count": high_priority_p1,
        "stale_10d_count": stale_10d_count,
        "external_attention_requests": external_requests,
        "reference_date": ref_ts.strftime("%Y-%m-%d"),
    }

    if df.empty:
        matrix_df = pd.DataFrame(columns=["DIVISION_UNIT", "P1", "P2", "P3", "TOTAL"])
    else:
        ct = pd.crosstab(df["DIVISION_UNIT"], df["PRIORITY"])
        ct = ct.reindex(columns=["1", "2", "3"], fill_value=0).rename(
            columns={"1": "P1", "2": "P2", "3": "P3"}
        )
        ct["TOTAL"] = ct["P1"] + ct["P2"] + ct["P3"]
        matrix_df = (
            ct.reset_index()
            .sort_values(by=["TOTAL", "DIVISION_UNIT"], ascending=[False, True])
            .reset_index(drop=True)
        )

    return kpis, matrix_df


def export_interactive_html_hub(
    df: pd.DataFrame,
    kpis: dict[str, object],
    matrix_df: pd.DataFrame,
    output_html_path: str | Path,
    output_json_path: str | Path | None = None,
) -> Path:
    """Serialize sanitized single-table records into JSON and a standalone interactive HTML Hub."""
    export_df = df.copy()
    export_df["LAST_INTERACTION_STR"] = (
        pd.to_datetime(export_df["LAST_INTERACTION_DATE"], errors="coerce")
        .dt.strftime("%Y-%m-%d")
        .fillna("")
    )

    records = [
        {
            "caseNumber": str(row.CASE_NUMBER),
            "theme": str(row.THEME),
            "synthesis": str(row.SYNTHESIS),
            "priority": str(row.PRIORITY),
            "division": str(row.DIVISION_UNIT),
            "status": str(row.ANALYSIS_STATUS),
            "internalOwner": str(row.INTERNAL_OWNER),
            "externalContact": str(row.EXTERNAL_CONTACT),
            "escalatedToBoard": bool(row.ESCALATED_TO_BOARD),
            "attentionFlag": bool(row.EXTERNAL_ATTENTION_FLAG),
            "latestInteraction": str(row.LATEST_INTERACTION),
            "lastInteractionDate": str(row.LAST_INTERACTION_STR),
            "eventsLog": str(row.EVENTS_LOG),
            "interactionsLog": str(row.INTERACTIONS_LOG),
            "caseLink": str(row.CASE_LINK),
        }
        for row in export_df.itertuples(index=False)
    ]

    if output_json_path is not None:
        json_out = Path(output_json_path)
        json_out.parent.mkdir(parents=True, exist_ok=True)
        json_out.write_text(
            json.dumps({"kpis": kpis, "matrix": matrix_df.to_dict(orient="records"), "cases": records}, indent=2),
            encoding="utf-8",
        )

    matrix_rows_html = "\n".join(
        f"<tr><td><strong>{escape(str(r.DIVISION_UNIT))}</strong></td>"
        f"<td>{int(r.P1)}</td><td>{int(r.P2)}</td><td>{int(r.P3)}</td>"
        f"<td><strong>{int(r.TOTAL)}</strong></td></tr>"
        for r in matrix_df.itertuples(index=False)
    )

    embedded_json = json.dumps(records, ensure_ascii=False)
    html_doc = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Enterprise Operations &amp; Compliance Intelligence Hub</title>
  <style>
    :root {{ --bg:#F5F6FA; --card:#FFFFFF; --accent:#5B2C83; --text:#1F2430; --muted:#64748B; --border:#E2E8F0; }}
    * {{ box-sizing:border-box; }}
    body {{ margin:0; font-family:Inter,"Segoe UI",Arial,sans-serif; background:var(--bg); color:var(--text); font-size:14px; }}
    header {{ background:linear-gradient(120deg,#3B185F 0%,#5B2C83 100%); color:#fff; padding:16px 28px; display:flex; justify-content:space-between; align-items:center; }}
    header h1 {{ margin:0; font-size:18px; }}
    main {{ max-width:1440px; margin:0 auto; padding:20px 28px; }}
    .kpis {{ display:grid; grid-template-columns:repeat(4,1fr); gap:14px; margin-bottom:18px; }}
    .card {{ background:var(--card); border:1px solid var(--border); border-radius:10px; padding:16px; }}
    .kpi-title {{ color:var(--muted); font-size:12px; font-weight:600; }}
    .kpi-val {{ font-size:28px; font-weight:700; margin:4px 0; }}
    .kpi-sub {{ color:var(--muted); font-size:11px; }}
    .toolbar {{ display:flex; gap:10px; align-items:center; margin-bottom:14px; flex-wrap:wrap; }}
    input, select, button {{ padding:8px 12px; border:1px solid var(--border); border-radius:8px; font:inherit; }}
    button.legend-btn {{ background:#F3E8FF; color:var(--accent); font-weight:600; cursor:pointer; }}
    .popover {{ display:none; background:#fff; border:1px solid var(--accent); border-radius:8px; padding:10px 14px; margin-bottom:12px; font-size:12px; }}
    .popover.open {{ display:block; }}
    table {{ width:100%; border-collapse:collapse; background:#fff; border:1px solid var(--border); border-radius:10px; overflow:hidden; }}
    th, td {{ text-align:left; padding:10px 12px; border-bottom:1px solid var(--border); vertical-align:top; }}
    th {{ background:#F8FAFC; font-size:11px; text-transform:uppercase; color:var(--muted); }}
    tr.closed-row {{ background:#F8FAFC; color:var(--muted); }}
    tr:hover {{ background:#F3E8FF44; cursor:pointer; }}
    .badge {{ display:inline-block; padding:2px 7px; border-radius:6px; font-size:11px; font-weight:600; }}
    .badge-closed {{ background:#E2E8F0; color:#475569; margin-right:6px; }}
    .badge-board {{ background:#FEF3C7; color:#92400E; margin-left:6px; }}
    .drawer {{ position:fixed; top:0; right:-520px; width:min(500px,100vw); height:100vh; background:#fff; box-shadow:-8px 0 28px rgba(15,23,42,.18); padding:22px; overflow-y:auto; transition:right .2s ease; z-index:20; }}
    .drawer.open {{ right:0; }}
    pre {{ background:#F8FAFC; border:1px solid var(--border); border-radius:8px; padding:10px; white-space:pre-wrap; font-family:inherit; font-size:12px; }}
  </style>
</head>
<body>
  <header>
    <h1>Enterprise Operations &amp; Compliance Intelligence Hub</h1>
    <span>Single-Table Dual-Stream View (Ref: {escape(str(kpis['reference_date']))})</span>
  </header>
  <main>
    <section class="kpis">
      <div class="card"><div class="kpi-title">Monitored Cases</div><div class="kpi-val">{kpis['total_monitored']}</div><div class="kpi-sub">{kpis['in_progress_count']} in progress &middot; {kpis['closed_count']} closed</div></div>
      <div class="card"><div class="kpi-title">High Priority (P1)</div><div class="kpi-val">{kpis['high_priority_p1_count']}</div><div class="kpi-sub">Immediate executive attention</div></div>
      <div class="card"><div class="kpi-title">Stale &ge; 10 Days</div><div class="kpi-val">{kpis['stale_10d_count']}</div><div class="kpi-sub">Active cases only (Closed excluded)</div></div>
      <div class="card"><div class="kpi-title">Attention Requests</div><div class="kpi-val">{kpis['external_attention_requests']}</div><div class="kpi-sub">External team escalations</div></div>
    </section>

    <section class="card" style="margin-bottom:18px;">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
        <strong>Priority Matrix by Division</strong>
        <button class="legend-btn" onclick="document.getElementById('legendPop').classList.toggle('open')">&#9432; Priority Legend</button>
      </div>
      <div id="legendPop" class="popover">1 = High Priority (Immediate Attention) &middot; 2 = Medium Priority (Regular Tracking) &middot; 3 = Low Priority (Monitoring)</div>
      <table><thead><tr><th>Division</th><th>1</th><th>2</th><th>3</th><th>Total</th></tr></thead><tbody>{matrix_rows_html}</tbody></table>
    </section>

    <section>
      <div class="toolbar">
        <input id="searchBox" placeholder="Search Case ID or Theme..." oninput="renderTable()" style="min-width:240px;" />
        <select id="statusFilter" onchange="renderTable()"><option value="ALL">All Statuses</option><option value="In Progress">In Progress</option><option value="Closed">Closed</option></select>
        <select id="prioFilter" onchange="renderTable()"><option value="ALL">All Priorities</option><option value="1">P1</option><option value="2">P2</option><option value="3">P3</option></select>
      </div>
      <table>
        <thead><tr><th>Theme / Synthesis</th><th>Case ID</th><th>Prior.</th><th>Owner</th><th>Division</th><th>Latest History (Interactions Only)</th></tr></thead>
        <tbody id="casesBody"></tbody>
      </table>
    </section>
  </main>

  <aside id="xrayDrawer" class="drawer">
    <div style="display:flex;justify-content:space-between;align-items:center;">
      <h2 id="drawerTitle" style="margin:0;font-size:17px;">Case X-Ray</h2>
      <button onclick="document.getElementById('xrayDrawer').classList.remove('open')">Close</button>
    </div>
    <p id="drawerMeta" style="color:#64748B;font-size:12px;"></p>
    <h4>5. Process Events (Audit Trail [[EVENTS]])</h4>
    <pre id="drawerEvents"></pre>
    <h4>6. Executive Interactions ([[INTERACTIONS]])</h4>
    <pre id="drawerInteractions"></pre>
  </aside>

  <script>
    const CASES = {embedded_json};
    function renderTable() {{
      const q = document.getElementById("searchBox").value.toLowerCase();
      const st = document.getElementById("statusFilter").value;
      const pr = document.getElementById("prioFilter").value;
      const filtered = CASES.filter(c =>
        (st === "ALL" || c.status === st) &&
        (pr === "ALL" || c.priority === pr) &&
        (!q || c.caseNumber.toLowerCase().includes(q) || c.theme.toLowerCase().includes(q))
      );
      document.getElementById("casesBody").innerHTML = filtered.map((c, idx) => `
        <tr class="${{c.status === 'Closed' ? 'closed-row' : ''}}" onclick="openDrawer(${{CASES.indexOf(c)}})">
          <td>${{c.status === 'Closed' ? '<span class="badge badge-closed">[Closed]</span>' : ''}}<strong>${{c.theme}}</strong>${{c.escalatedToBoard ? '<span class="badge badge-board">&uarr; BOARD</span>' : ''}}<br/><small>${{c.synthesis}}</small></td>
          <td>${{c.caseNumber}}</td>
          <td>${{c.priority}}</td>
          <td>${{c.internalOwner}}</td>
          <td>${{c.division}}</td>
          <td>${{c.latestInteraction || '<span style="color:#94A3B8">No executive interaction recorded</span>'}}</td>
        </tr>`).join("");
    }}
    function openDrawer(idx) {{
      const c = CASES[idx];
      document.getElementById("drawerTitle").textContent = c.theme;
      document.getElementById("drawerMeta").textContent = `${{c.caseNumber}} | ${{c.division}} | Priority ${{c.priority}} | ${{c.status}}`;
      document.getElementById("drawerEvents").textContent = c.eventsLog || "No system events recorded.";
      document.getElementById("drawerInteractions").textContent = c.interactionsLog || "No executive interactions recorded.";
      document.getElementById("xrayDrawer").classList.add("open");
    }}
    renderTable();
  </script>
</body>
</html>"""

    out_html = Path(output_html_path)
    out_html.parent.mkdir(parents=True, exist_ok=True)
    out_html.write_text(html_doc, encoding="utf-8")
    return out_html


def run_operational_pipeline(
    input_path: str | Path,
    output_html_path: str | Path,
    output_json_path: str | Path | None = None,
    sheet_name: str | int = 0,
    reference_date: str | pd.Timestamp | None = None,
) -> dict[str, object]:
    """Run Single-Table Operational ETL and emit interactive HTML Hub + JSON."""
    path = Path(input_path)
    raw_df = pd.read_csv(path) if path.suffix.lower() == ".csv" else pd.read_excel(path, sheet_name=sheet_name)
    sanitized_df = sanitize_single_table_processes(raw_df)
    kpis, matrix_df = compute_executive_kpis_and_matrix(sanitized_df, reference_date=reference_date)
    html_path = export_interactive_html_hub(
        sanitized_df,
        kpis,
        matrix_df,
        output_html_path=output_html_path,
        output_json_path=output_json_path,
    )
    return {
        "output_html": str(html_path),
        "kpis": kpis,
        "divisions": int(len(matrix_df)),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Single-Table Operational Data Pipeline & HTML Hub")
    parser.add_argument("--input", required=True, help="Single-table tracker workbook (.xlsx or .csv)")
    parser.add_argument(
        "--output-html",
        default="data/output/operations_intelligence_hub.html",
        help="Destination HTML dashboard path",
    )
    parser.add_argument(
        "--output-json",
        default="data/output/single_table_cases.json",
        help="Optional JSON export path",
    )
    parser.add_argument("--ref-date", default=None, help="Optional YYYY-MM-DD reference date for stale SLA check")
    args = parser.parse_args()

    res = run_operational_pipeline(
        input_path=args.input,
        output_html_path=args.output_html,
        output_json_path=args.output_json,
        reference_date=args.ref_date,
    )
    print(f"HTML Hub exported -> {res['output_html']} (KPIs: {res['kpis']})")


if __name__ == "__main__":
    main()
