import streamlit as st
import json
import zipfile
import re
import io
import os
import base64
import requests
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

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;700&family=Inter:wght@300;400;500;600&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.title-block {
    background: linear-gradient(135deg, #0f1729 0%, #1a2440 100%);
    border: 1px solid #2a3a5c;
    border-radius: 12px;
    padding: 2rem 2.5rem;
    margin-bottom: 1.5rem;
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
.title-block h1 { font-family: 'JetBrains Mono', monospace; color: #e2e8f0; font-size: 1.8rem; margin: 0 0 0.3rem 0; }
.title-block p { color: #64748b; margin: 0; font-size: 0.9rem; }
.metric-card { background: #0f1729; border: 1px solid #1e2d47; border-radius: 10px; padding: 1.2rem 1.5rem; text-align: center; }
.metric-card .value { font-family: 'JetBrains Mono', monospace; font-size: 2.2rem; font-weight: 700; line-height: 1; margin-bottom: 0.4rem; }
.metric-card .label { color: #64748b; font-size: 0.78rem; text-transform: uppercase; letter-spacing: 0.08em; }
.critical .value { color: #ef4444; }
.warning  .value { color: #f59e0b; }
.info     .value { color: #3b82f6; }
.success  .value { color: #10b981; }
</style>
""", unsafe_allow_html=True)

# ─── GitHub Storage ───────────────────────────────────────────────────────────
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")
GITHUB_REPO  = os.environ.get("GITHUB_REPO", "")
DATA_FILE    = "data/gb300_data.json"

def gh_headers():
    return {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}

def load_from_github():
    if not GITHUB_TOKEN or not GITHUB_REPO:
        return []
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{DATA_FILE}"
    r = requests.get(url, headers=gh_headers())
    if r.status_code == 200:
        content = base64.b64decode(r.json()["content"]).decode("utf-8")
        return json.loads(content)
    return []

def save_to_github(data):
    if not GITHUB_TOKEN or not GITHUB_REPO:
        return False
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{DATA_FILE}"
    # Get current SHA if file exists
    r = requests.get(url, headers=gh_headers())
    sha = r.json().get("sha") if r.status_code == 200 else None
    content = base64.b64encode(json.dumps(data, ensure_ascii=False, indent=2).encode()).decode()
    payload = {
        "message": f"Update GB300 data — {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "content": content,
    }
    if sha:
        payload["sha"] = sha
    r = requests.put(url, headers=gh_headers(), json=payload)
    return r.status_code in (200, 201)

# ─── Parser ───────────────────────────────────────────────────────────────────
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
        if not gf and subclass2 == "62660":
            errors.append("GPU missing from nvidia-smi")

    for gf in gpu_faults:
        errors.append(f"GPU BDF {gf['bdf']} missing — {gf['board']} Board")

    if not errors:
        errors.append(reason or f"FaultCode {fault_code}")

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

def parse_description_content(content, sn, folder_info):
    content = content.strip()
    if not content.startswith("{"):
        return None
    try:
        diag = json.loads(content)
    except Exception:
        return None
    errors, fix, component, source, gpu_faults, severity = classify_fault(diag)
    part = diag.get("PartFailures", [{}])[0]
    return {
        "SN":            sn,
        "Ticket":        folder_info.get("Ticket", ""),
        "Wal":           folder_info.get("Wal", ""),
        "Model":         part.get("ModelNumber", ""),
        "Location":      part.get("Location", ""),
        "FaultCode":     str(diag.get("FaultCode", "")),
        "Source":        source,
        "Timestamp":     part.get("DateandTimestamp", ""),
        "GPU_Total":     len(gpu_faults),
        "GPU_Primary":   sum(1 for g in gpu_faults if g["board"] == "Primary"),
        "GPU_Secondary": sum(1 for g in gpu_faults if g["board"] == "Secondary"),
        "Errors":        "\n".join(errors),
        "Action":        fix,
        "Component":     component,
        "Severity":      severity,
        "UploadedAt":    datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

def process_zip(zip_bytes):
    rows = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
        for name in z.namelist():
            parts = Path(name).parts
            if len(parts) < 3 or name.endswith("/"):
                continue
            sn, folder, filename = parts[0], parts[1], parts[-1]
            if "description" not in filename.lower() and not filename.endswith(".json"):
                continue
            try:
                content = z.read(name).decode("utf-8", errors="ignore")
                folder_info = parse_folder_name(folder)
                row = parse_description_content(content, sn, folder_info)
                if row:
                    rows.append(row)
            except Exception:
                continue
    return rows

def process_description_files(files):
    """Process individually uploaded description files."""
    rows = []
    for f in files:
        try:
            content = f.read().decode("utf-8", errors="ignore")
            # Try to extract SN from filename or content
            row = parse_description_content(content, f.name, {})
            if row:
                # Try get SN from JSON
                try:
                    diag = json.loads(content.strip())
                    pf = diag.get("PartFailures", [{}])[0]
                    sn = pf.get("SerialNumber", f.name)
                    row["SN"] = sn
                except Exception:
                    pass
                rows.append(row)
        except Exception:
            continue
    return rows

def build_excel(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "GB300 Analysis"
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    header_fill = PatternFill("solid", start_color="1F3864")
    cell_font   = Font(name="Arial", size=9)
    center      = Alignment(horizontal="center", vertical="top", wrap_text=True)
    left        = Alignment(horizontal="left",   vertical="top", wrap_text=True)
    thin        = Side(style="thin", color="CCCCCC")
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)
    fills       = {"critical": "FFCCCC", "warning": "FFE5CC", "low": "FFFACC", "info": "FFFFFF"}

    headers   = ["SN","Ticket","Wal","Model","Location","Fault Code","Source","Timestamp","GPUs Missing","Primary","Secondary","Errors","Action Required","Component","Uploaded At"]
    col_keys  = ["SN","Ticket","Wal","Model","Location","FaultCode","Source","Timestamp","GPU_Total","GPU_Primary","GPU_Secondary","Errors","Action","Component","UploadedAt"]
    col_widths= [22,14,18,14,10,10,12,20,12,10,11,55,40,18,18]

    for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font = header_font; cell.fill = header_fill
        cell.alignment = center; cell.border = border
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    for ri, row in enumerate(rows, 2):
        rf = PatternFill("solid", start_color=fills.get(row["Severity"], "FFFFFF"))
        for ci, key in enumerate(col_keys, 1):
            cell = ws.cell(row=ri, column=ci, value=row.get(key, ""))
            cell.font = cell_font; cell.fill = rf; cell.border = border
            cell.alignment = center if key in ("SN","FaultCode","GPU_Total","GPU_Primary","GPU_Secondary","Location") else left
        ws.row_dimensions[ri].height = max(18, 14 * (row["Errors"].count("\n") + 1))

    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "GB300 Analysis Summary"
    ws2["A1"].font = Font(name="Arial", bold=True, size=14, color="1F3864")
    ws2["B1"] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws2["B1"].font = Font(name="Arial", size=10, color="888888")
    for i, (label, val) in enumerate([
        ("Total Servers", len(rows)),
        ("Critical (4+ GPU / PS fault)", sum(1 for r in rows if r["Severity"] == "critical")),
        ("Warning (2-3 GPU)", sum(1 for r in rows if r["Severity"] == "warning")),
        ("Low (1 GPU)", sum(1 for r in rows if r["Severity"] == "low")),
        ("PS RUN PWR FAULT", sum(1 for r in rows if "PS RUN PWR FAULT" in r["Errors"])),
        ("Total GPUs Missing", sum(r["GPU_Total"] for r in rows)),
    ], 3):
        ws2.cell(row=i, column=1, value=label).font = Font(name="Arial", bold=True, size=11)
        ws2.cell(row=i, column=2, value=val).font   = Font(name="Arial", size=11)
    ws2.column_dimensions["A"].width = 35
    ws2.column_dimensions["B"].width = 15

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ─── Session state ────────────────────────────────────────────────────────────
if "all_rows" not in st.session_state:
    with st.spinner("Loading data from GitHub..."):
        st.session_state.all_rows = load_from_github()
if "delete_mode" not in st.session_state:
    st.session_state.delete_mode = False

# ─── Header ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="title-block">
  <h1>🖥️ GB300 Log Analyzer</h1>
  <p>Quanta Azure GPU Compute · B200 · CSI Diagnostic Parser — Data saved permanently</p>
</div>
""", unsafe_allow_html=True)

# ─── Search bar (prominent, always visible) ───────────────────────────────────
search = st.text_input("🔍 Search by SN, Ticket, Wal, or Error", placeholder="e.g. P87815452... or PS RUN or 62660", label_visibility="visible")

st.markdown("---")

# ─── Upload section ───────────────────────────────────────────────────────────
with st.expander("📂 Upload New Logs", expanded=not bool(st.session_state.all_rows)):
    tab1, tab2 = st.tabs(["📦 Upload ZIP (nhiều SN)", "📄 Upload Description file (1 SN)"])

    with tab1:
        uploaded_zip = st.file_uploader("Chọn GB300Logs.zip", type=["zip"], accept_multiple_files=True, key="zip_uploader")
        if uploaded_zip:
            if st.button("➕ Process & Save ZIP", type="primary"):
                new_rows = []
                for f in uploaded_zip:
                    with st.spinner(f"Processing {f.name}..."):
                        new_rows.extend(process_zip(f.read()))
                existing_sns = {r["SN"] for r in st.session_state.all_rows}
                added = [r for r in new_rows if r["SN"] not in existing_sns]
                if added:
                    st.session_state.all_rows.extend(added)
                    with st.spinner("Saving to GitHub..."):
                        ok = save_to_github(st.session_state.all_rows)
                    if ok:
                        st.success(f"✅ Added {len(added)} new servers — saved permanently!")
                    else:
                        st.warning(f"✅ Added {len(added)} servers (session only — GitHub save failed)")
                else:
                    st.info("No new SNs found (already in database)")

    with tab2:
        uploaded_files = st.file_uploader("Chọn Description file(s)", accept_multiple_files=True, key="file_uploader")
        if uploaded_files:
            if st.button("➕ Process & Save Files", type="primary"):
                new_rows = process_description_files(uploaded_files)
                existing_sns = {r["SN"] for r in st.session_state.all_rows}
                added = [r for r in new_rows if r["SN"] not in existing_sns]
                if added:
                    st.session_state.all_rows.extend(added)
                    with st.spinner("Saving to GitHub..."):
                        ok = save_to_github(st.session_state.all_rows)
                    if ok:
                        st.success(f"✅ Added {len(added)} new servers — saved permanently!")
                    else:
                        st.warning(f"✅ Added (session only — GitHub save failed)")
                else:
                    st.info("No new SNs found")

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
    fc1, fc2 = st.columns(2)
    with fc1:
        sev_filter = st.multiselect("Severity", ["critical","warning","low","info"],
                                     default=["critical","warning","low","info"])
    with fc2:
        sources = list(set(r["Source"] for r in rows))
        source_filter = st.multiselect("Source", sources, default=sources)

    # Apply filters + search
    filtered = [
        r for r in rows
        if r["Severity"] in sev_filter
        and r["Source"] in source_filter
        and (not search or any(
            search.lower() in str(r.get(k, "")).lower()
            for k in ["SN","Ticket","Wal","Errors","Action","Model"]
        ))
    ]

    # ─── Table + Delete ──────────────────────────────────────────────────────
    col_title, col_del_btn = st.columns([4, 1])
    with col_title:
        st.markdown(f"### Results — {len(filtered)} servers")
    with col_del_btn:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🗑️ Delete Mode" if not st.session_state.delete_mode else "✅ Exit Delete", use_container_width=True):
            st.session_state.delete_mode = not st.session_state.delete_mode
            st.rerun()

    if st.session_state.delete_mode:
        st.warning("⚠️ Delete mode ON — chọn SN muốn xóa bên dưới")
        sns_to_delete = []
        cols = st.columns(4)
        for i, r in enumerate(filtered):
            with cols[i % 4]:
                if st.checkbox(f"{r['SN']}", key=f"del_{r['SN']}"):
                    sns_to_delete.append(r["SN"])

        if sns_to_delete:
            dc1, dc2 = st.columns(2)
            with dc1:
                st.error(f"Sẽ xóa {len(sns_to_delete)} SN: {', '.join(sns_to_delete)}")
            with dc2:
                if st.button("🗑️ Confirm Delete", type="primary", use_container_width=True):
                    st.session_state.all_rows = [r for r in st.session_state.all_rows if r["SN"] not in sns_to_delete]
                    with st.spinner("Saving..."):
                        save_to_github(st.session_state.all_rows)
                    st.session_state.delete_mode = False
                    st.success(f"✅ Deleted {len(sns_to_delete)} servers")
                    st.rerun()
    else:
        # Normal table view
        df = pd.DataFrame(filtered)[["SN","Ticket","Wal","Model","FaultCode","Source","Timestamp","GPU_Total","GPU_Primary","GPU_Secondary","Errors","Action","Severity","UploadedAt"]]
        df.columns = ["SN","Ticket","Wal","Model","Fault Code","Source","Timestamp","GPUs Missing","Primary","Secondary","Errors","Action Required","Severity","Uploaded At"]

        def color_severity(val):
            colors = {
                "critical": "background-color:#3d1515;color:#ef4444",
                "warning":  "background-color:#3d2a10;color:#f59e0b",
                "low":      "background-color:#2a2d10;color:#eab308",
                "info":     "background-color:#0f1729;color:#64748b"
            }
            return colors.get(val, "")

        styled = df.style.map(color_severity, subset=["Severity"])
        st.dataframe(styled, use_container_width=True, height=500, hide_index=True)

    # ─── Download ────────────────────────────────────────────────────────────
    st.markdown("---")
    dl1, dl2 = st.columns(2)
    fname = f"GB300_Analysis_{datetime.now().strftime('%Y%m%d_%H%M')}"
    with dl1:
        excel_bytes = build_excel(filtered)
        st.download_button("⬇️ Download Excel", data=excel_bytes,
                           file_name=f"{fname}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
    with dl2:
        csv = df.to_csv(index=False).encode("utf-8") if not st.session_state.delete_mode else pd.DataFrame(filtered).to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download CSV", data=csv,
                           file_name=f"{fname}.csv",
                           mime="text/csv", use_container_width=True)

else:
    st.markdown("""
    <div style="text-align:center;padding:4rem;color:#334155;">
        <div style="font-size:3rem;margin-bottom:1rem">📂</div>
        <div style="font-family:'JetBrains Mono',monospace;font-size:1rem">No data yet — upload GB300Logs.zip to begin</div>
    </div>
    """, unsafe_allow_html=True)
