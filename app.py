import streamlit as st
import json
import zipfile
import re
import io
import tempfile
import os
from pathlib import Path
from datetime import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="GB300 Log Analyzer",
    page_icon="🖥️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── CSS ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;700&family=Inter:wght@300;400;500;600&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

.main { background: #0a0e1a; }
.block-container { padding: 2rem 3rem; max-width: 1400px; }

.title-block {
    background: linear-gradient(135deg, #0f1729 0%, #1a2440 100%);
    border: 1px solid #2a3a5c;
    border-radius: 12px;
    padding: 2rem 2.5rem;
    margin-bottom: 2rem;
    position: relative;
    overflow: hidden;
}
.title-block::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
    background: linear-gradient(90deg, #3b82f6, #06b6d4, #3b82f6);
}
.title-block h1 {
    font-family: 'JetBrains Mono', monospace;
    color: #e2e8f0;
    font-size: 1.8rem;
    margin: 0 0 0.3rem 0;
    letter-spacing: -0.5px;
}
.title-block p { color: #64748b; margin: 0; font-size: 0.9rem; }

.metric-card {
    background: #0f1729;
    border: 1px solid #1e2d47;
    border-radius: 10px;
    padding: 1.2rem 1.5rem;
    text-align: center;
}
.metric-card .value {
    font-family: 'JetBrains Mono', monospace;
    font-size: 2.2rem;
    font-weight: 700;
    line-height: 1;
    margin-bottom: 0.4rem;
}
.metric-card .label { color: #64748b; font-size: 0.78rem; text-transform: uppercase; letter-spacing: 0.08em; }

.critical .value { color: #ef4444; }
.warning  .value { color: #f59e0b; }
.info     .value { color: #3b82f6; }
.success  .value { color: #10b981; }

.upload-zone {
    background: #0f1729;
    border: 2px dashed #2a3a5c;
    border-radius: 12px;
    padding: 2rem;
    text-align: center;
    margin-bottom: 1.5rem;
    transition: border-color 0.2s;
}
.upload-zone:hover { border-color: #3b82f6; }

.sn-badge {
    display: inline-block;
    background: #1e2d47;
    color: #94a3b8;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.72rem;
    padding: 2px 8px;
    border-radius: 4px;
    margin: 2px;
}

.fault-critical { color: #ef4444; font-weight: 600; }
.fault-warning  { color: #f59e0b; font-weight: 600; }
.fault-low      { color: #3b82f6; }

stDataFrame { font-family: 'JetBrains Mono', monospace !important; }
</style>
""", unsafe_allow_html=True)

# ─── Parser logic ────────────────────────────────────────────────────────────
FAULT_DESCRIPTIONS = {
    "62660": "GPU missing from nvidia-smi",
    "67802": "Server Chassis Fault",
}

def parse_folder_name(folder_name):
    info = {}
    wal_match = re.search(r"(WAL_[^\-\(\)/\\]+)", folder_name)
    if wal_match:
        info["Wal"] = wal_match.group(1)
    for key, pattern in [
        ("Mfr",    r"\(Mfr\)-([^-\(]+)"),
        ("SI",     r"\(SI\)-([^-\(]+)"),
        ("Ticket", r"\(Ticket\)-([^-\(]+)"),
        ("SN",     r"\(SN\)-([^-\(]+)"),
    ]:
        m = re.search(pattern, folder_name)
        if m:
            info[key] = m.group(1).strip()
    return info

def parse_gpu_faults(diagnostic_summary):
    faults = []
    for entry in diagnostic_summary.split("/"):
        bdf_match = re.search(r"GPU with BDF ([\w\-\.]+) is missing", entry)
        loc_match = re.search(r"PhysicalLocation:([^;]+)", entry)
        if bdf_match:
            bdf = bdf_match.group(1)
            loc = loc_match.group(1).strip() if loc_match else ""
            board = "Primary" if "Primary" in loc else ("Secondary" if "Secondary" in loc else "")
            faults.append({"bdf": bdf, "location": loc, "board": board})
    return faults

def classify_fault(diag):
    errors = []
    reason = diag.get("Reason", "")
    component = diag.get("ComponentType", "Unknown")
    source = diag.get("SourceOfFault", "")
    fault_code = str(diag.get("FaultCode", ""))
    action = diag.get("Action", "Unknown")

    is_ps_fault = "PS RUN PWR FAULT" in reason
    if is_ps_fault:
        errors.append("PS RUN PWR FAULT")
    if "Vendor RMA" in reason:
        errors.append("Vendor RMA Required")

    gpu_faults = []
    for pf in diag.get("PartFailures", []):
        summary = pf.get("Diagnostic_Summary", "")
        subclass2 = str(pf.get("Subclass2", ""))
        gf = parse_gpu_faults(summary)
        gpu_faults.extend(gf)
        if not gf and subclass2 in FAULT_DESCRIPTIONS:
            errors.append(FAULT_DESCRIPTIONS[subclass2])

    for gf in gpu_faults:
        errors.append(f"GPU BDF {gf['bdf']} missing — {gf['board']} Board")

    if not errors:
        errors.append(reason or f"FaultCode {fault_code}")

    # Fix action
    if is_ps_fault:
        fix = "Replace Power Supply (PS)"
    elif gpu_faults:
        boards = sorted(set(gf["board"] for gf in gpu_faults if gf["board"]))
        fix = "Replace " + (" & ".join(boards) + " GPU Board" if boards else "GPU Assembly")
    else:
        fix = f"Replace {component}"

    severity = "critical" if (is_ps_fault or len(gpu_faults) >= 4) else \
               "warning" if len(gpu_faults) >= 2 else \
               "low" if len(gpu_faults) == 1 else "info"

    return errors, fix, component, source, gpu_faults, severity

def process_zip(zip_bytes):
    rows = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
        names = z.namelist()
        for name in names:
            parts = Path(name).parts
            if len(parts) < 3:
                continue
            sn = parts[0]
            folder = parts[1]
            filename = parts[-1]

            if len(parts) < 3 or parts[-1] == "":
                continue
            if "description" not in filename.lower() and not filename.endswith(".json"):
                # try to read anyway if it looks like a file (not dir)
                if not filename or filename.endswith("/"):
                    continue

            try:
                content = z.read(name).decode("utf-8", errors="ignore").strip()
                if not content.startswith("{"):
                    continue
                diag = json.loads(content)
            except Exception:
                continue

            folder_info = parse_folder_name(folder)
            errors, fix, component, source, gpu_faults, severity = classify_fault(diag)

            part = diag.get("PartFailures", [{}])[0]
            timestamp = part.get("DateandTimestamp", "")
            model = part.get("ModelNumber", "")
            location = part.get("Location", "")

            primary_count   = sum(1 for g in gpu_faults if g["board"] == "Primary")
            secondary_count = sum(1 for g in gpu_faults if g["board"] == "Secondary")

            rows.append({
                "SN":              sn,
                "Ticket":          folder_info.get("Ticket", ""),
                "Wal":             folder_info.get("Wal", ""),
                "Model":           model,
                "Location":        location,
                "FaultCode":       str(diag.get("FaultCode", "")),
                "Source":          source,
                "Timestamp":       timestamp,
                "GPU_Total":       len(gpu_faults),
                "GPU_Primary":     primary_count,
                "GPU_Secondary":   secondary_count,
                "Errors":          "\n".join(errors),
                "Action":          fix,
                "Component":       component,
                "Severity":        severity,
                "RawReason":       diag.get("Reason", ""),
            })
    return rows

def build_excel(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "GB300 Analysis"

    header_font  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    header_fill  = PatternFill("solid", start_color="1F3864")
    cell_font    = Font(name="Arial", size=9)
    center       = Alignment(horizontal="center", vertical="top", wrap_text=True)
    left         = Alignment(horizontal="left",   vertical="top", wrap_text=True)
    thin         = Side(style="thin", color="CCCCCC")
    border       = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill_red     = PatternFill("solid", start_color="FFCCCC")
    fill_orange  = PatternFill("solid", start_color="FFE5CC")
    fill_yellow  = PatternFill("solid", start_color="FFFACC")
    fill_white   = PatternFill("solid", start_color="FFFFFF")

    headers   = ["SN","Ticket","Wal","Model","Location","Fault Code","Source","Timestamp","GPUs Missing","Primary","Secondary","Errors","Action Required","Component"]
    col_keys  = ["SN","Ticket","Wal","Model","Location","FaultCode","Source","Timestamp","GPU_Total","GPU_Primary","GPU_Secondary","Errors","Action","Component"]
    col_widths= [22,14,18,14,10,10,12,20,12,10,11,55,40,18]

    for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font = header_font; cell.fill = header_fill
        cell.alignment = center; cell.border = border
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    for ri, row in enumerate(rows, 2):
        sev = row["Severity"]
        rf = fill_red if sev == "critical" else fill_orange if sev == "warning" else fill_yellow if sev == "low" else fill_white
        for ci, key in enumerate(col_keys, 1):
            cell = ws.cell(row=ri, column=ci, value=row.get(key, ""))
            cell.font = cell_font; cell.fill = rf; cell.border = border
            cell.alignment = center if key in ("SN","FaultCode","GPU_Total","GPU_Primary","GPU_Secondary","Location") else left
        lines = row["Errors"].count("\n") + 1
        ws.row_dimensions[ri].height = max(18, 14 * lines)

    # Summary sheet
    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "GB300 Log Analysis Summary"
    ws2["A1"].font = Font(name="Arial", bold=True, size=14, color="1F3864")
    ws2["B1"] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws2["B1"].font = Font(name="Arial", size=10, color="888888")

    summary = [
        ("Total Servers Analyzed",        len(rows)),
        ("🔴 Critical (4+ GPU / PS fault)", sum(1 for r in rows if r["Severity"] == "critical")),
        ("🟠 Warning (2-3 GPU missing)",   sum(1 for r in rows if r["Severity"] == "warning")),
        ("🟡 Low (1 GPU missing)",         sum(1 for r in rows if r["Severity"] == "low")),
        ("PS RUN PWR FAULT",              sum(1 for r in rows if "PS RUN PWR FAULT" in r["Errors"])),
        ("Total GPUs Missing",            sum(r["GPU_Total"] for r in rows)),
        ("Primary Board GPU Missing",     sum(r["GPU_Primary"] for r in rows)),
        ("Secondary Board GPU Missing",   sum(r["GPU_Secondary"] for r in rows)),
    ]
    for i, (label, val) in enumerate(summary, 3):
        ws2.cell(row=i, column=1, value=label).font = Font(name="Arial", bold=True, size=11)
        ws2.cell(row=i, column=2, value=val).font   = Font(name="Arial", size=11)
    ws2.column_dimensions["A"].width = 38
    ws2.column_dimensions["B"].width = 15

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ─── Session state ────────────────────────────────────────────────────────────
if "all_rows" not in st.session_state:
    st.session_state.all_rows = []
if "uploaded_files" not in st.session_state:
    st.session_state.uploaded_files = set()

# ─── Header ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="title-block">
  <h1>🖥️ GB300 Log Analyzer</h1>
  <p>Quanta Azure GPU Compute · B200 · CSI Diagnostic Parser</p>
</div>
""", unsafe_allow_html=True)

# ─── Upload ───────────────────────────────────────────────────────────────────
col_up, col_clear = st.columns([5, 1])
with col_up:
    uploaded = st.file_uploader(
        "Upload GB300Logs.zip (có thể upload nhiều lần để append data)",
        type=["zip"],
        accept_multiple_files=True,
        label_visibility="visible",
    )
with col_clear:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🗑️ Clear All", use_container_width=True):
        st.session_state.all_rows = []
        st.session_state.uploaded_files = set()
        st.rerun()

# Process new uploads
if uploaded:
    new_count = 0
    for f in uploaded:
        file_key = f"{f.name}_{f.size}"
        if file_key not in st.session_state.uploaded_files:
            with st.spinner(f"Processing {f.name}..."):
                new_rows = process_zip(f.read())
                # Avoid duplicate SNs
                existing_sns = {r["SN"] for r in st.session_state.all_rows}
                added = [r for r in new_rows if r["SN"] not in existing_sns]
                st.session_state.all_rows.extend(added)
                st.session_state.uploaded_files.add(file_key)
                new_count += len(added)
    if new_count > 0:
        st.success(f"✅ Added {new_count} new servers")

rows = st.session_state.all_rows

# ─── Metrics ─────────────────────────────────────────────────────────────────
if rows:
    total     = len(rows)
    critical  = sum(1 for r in rows if r["Severity"] == "critical")
    warning   = sum(1 for r in rows if r["Severity"] == "warning")
    gpu_total = sum(r["GPU_Total"] for r in rows)
    ps_faults = sum(1 for r in rows if "PS RUN PWR FAULT" in r["Errors"])

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.markdown(f'<div class="metric-card info"><div class="value">{total}</div><div class="label">Total Servers</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="metric-card critical"><div class="value">{critical}</div><div class="label">Critical</div></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="metric-card warning"><div class="value">{warning}</div><div class="label">Warning</div></div>', unsafe_allow_html=True)
    with c4:
        st.markdown(f'<div class="metric-card critical"><div class="value">{gpu_total}</div><div class="label">GPUs Missing</div></div>', unsafe_allow_html=True)
    with c5:
        st.markdown(f'<div class="metric-card warning"><div class="value">{ps_faults}</div><div class="label">PS Faults</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ─── Filters ─────────────────────────────────────────────────────────────
    with st.expander("🔍 Filters", expanded=False):
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            sev_filter = st.multiselect("Severity", ["critical","warning","low","info"], default=["critical","warning","low","info"])
        with fc2:
            source_filter = st.multiselect("Source", list(set(r["Source"] for r in rows)), default=list(set(r["Source"] for r in rows)))
        with fc3:
            search = st.text_input("Search SN / Ticket / Error", "")

    filtered = [
        r for r in rows
        if r["Severity"] in sev_filter
        and r["Source"] in source_filter
        and (not search or search.lower() in r["SN"].lower() or search.lower() in r["Ticket"].lower() or search.lower() in r["Errors"].lower())
    ]

    # ─── Table ───────────────────────────────────────────────────────────────
    st.markdown(f"### Results — {len(filtered)} servers")

    df = pd.DataFrame(filtered)[["SN","Ticket","Wal","Model","FaultCode","Source","Timestamp","GPU_Total","GPU_Primary","GPU_Secondary","Errors","Action","Severity"]]
    df.columns = ["SN","Ticket","Wal","Model","Fault Code","Source","Timestamp","GPUs Missing","Primary","Secondary","Errors","Action Required","Severity"]

    def color_severity(val):
        colors = {"critical":"background-color:#3d1515;color:#ef4444","warning":"background-color:#3d2a10;color:#f59e0b","low":"background-color:#2a2d10;color:#eab308","info":"background-color:#0f1729;color:#64748b"}
        return colors.get(val, "")

    styled = df.style.applymap(color_severity, subset=["Severity"])
    st.dataframe(styled, use_container_width=True, height=500, hide_index=True)

    # ─── Download ────────────────────────────────────────────────────────────
    st.markdown("---")
    dl1, dl2 = st.columns(2)
    with dl1:
        excel_bytes = build_excel(filtered)
        fname = f"GB300_Analysis_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button("⬇️ Download Excel", data=excel_bytes, file_name=fname,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
    with dl2:
        csv = df.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download CSV", data=csv,
                           file_name=f"GB300_Analysis_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                           mime="text/csv", use_container_width=True)
else:
    st.markdown("""
    <div style="text-align:center;padding:4rem;color:#334155;">
        <div style="font-size:3rem;margin-bottom:1rem">📂</div>
        <div style="font-family:'JetBrains Mono',monospace;font-size:1rem">Upload GB300Logs.zip to begin</div>
        <div style="font-size:0.8rem;margin-top:0.5rem;color:#1e2d47">Supports multiple uploads — data will be appended automatically</div>
    </div>
    """, unsafe_allow_html=True)
