"""Destination Change - stable Streamlit entry point.

Keeps the proven v5.0-style UI while avoiding pandas/backend imports during page bootstrap.
"""
from __future__ import annotations

import csv
import hashlib
import io
import os
import re
import tempfile
from datetime import date, timedelta
from pathlib import Path

import streamlit as st

st.set_page_config(page_title="Destination Change App", page_icon="📦", layout="wide")


def fmt_date(d: date) -> str:
    return f"{d.month}/{d.day}/{d.year}"


def saturday_of_current_week(today: date | None = None) -> date:
    if today is None:
        today = date.today()
    return today + timedelta(days=(5 - today.weekday()) % 7)


def normalize_whse(value) -> str:
    text = "" if value is None else str(value).strip()
    if not text:
        return ""
    try:
        f = float(text)
        if f.is_integer():
            return str(int(f))
    except Exception:
        pass
    return text.upper()


def normalize_vendor(value) -> str:
    text = "" if value is None else str(value).strip()
    text = text.strip('"')
    if re.fullmatch(r"0*\d+", text):
        return str(int(text))
    return text.upper()


def vendor_key(value) -> str:
    text = normalize_vendor(value)
    m = re.search(r"\((0*\d+)\)", text)
    if m:
        return str(int(m.group(1)))
    nums = re.findall(r"0*\d+", text)
    return str(int(nums[-1])) if nums else text


def _find_vendor_col(headers):
    exact = {str(h).strip().lower(): h for h in headers}
    for name in ["vendor", "vendor code", "vendorcode", "vendor #", "vendor#", "supplier", "supplier code"]:
        if name in exact:
            return exact[name]
    for h in headers:
        s = str(h).strip().lower()
        if "vendor" in s or "supplier" in s:
            return h
    return None


@st.cache_data(show_spinner=False, max_entries=8)
def detect_vendor_order_cached(payloads):
    rows = []
    seen = set()
    order = 1
    for file_order, (name, payload) in enumerate(payloads, 1):
        header = None
        data_rows = []
        try:
            text = payload.decode("utf-8-sig", errors="replace")
            reader = csv.reader(io.StringIO(text))
            for line in reader:
                if header is None and line and str(line[0]).strip().startswith("Item #"):
                    header = [str(x).strip() for x in line]
                    continue
                if header is not None and line:
                    data_rows.append(line)
        except Exception:
            continue
        if not header:
            continue
        vendor_col = _find_vendor_col(header)
        if vendor_col is None:
            continue
        try:
            vendor_idx = header.index(vendor_col)
        except ValueError:
            continue
        counts = {}
        for row in data_rows:
            if vendor_idx >= len(row):
                continue
            vk = vendor_key(row[vendor_idx])
            if not vk:
                continue
            counts[vk] = counts.get(vk, 0) + 1
            if vk not in seen:
                seen.add(vk)
                rows.append({"Vendor Order": order, "Vendor Code": vk, "Source PSW File Order": file_order, "Rows": 0})
                order += 1
        for r in rows:
            if r["Source PSW File Order"] == file_order and r["Vendor Code"] in counts:
                r["Rows"] = counts[r["Vendor Code"]]
    return rows


def save_uploaded(uploaded_file, folder: str, fallback_name: str) -> str:
    name = Path(uploaded_file.name or fallback_name).name
    stem = Path(name).stem.replace(" ", "_")
    suffix = Path(name).suffix or Path(fallback_name).suffix
    path = Path(folder) / f"{stem}{suffix}"
    counter = 1
    while path.exists():
        path = Path(folder) / f"{stem}_{counter}{suffix}"
        counter += 1
    path.write_bytes(uploaded_file.getbuffer())
    return str(path)


def render_logic_summary():
    with st.expander("Logic summary", expanded=False):
        st.markdown("""
**Inputs**
- PlanDetailTimeline.csv
- PSW / Production Schedule.csv
- One or more DueDateCalc.xlsx files

**Optimization**
- Main vendor is optimized first; sub-vendor uses Main Vendor After as baseline.
- Healthy SS% is checked across eligible Buy warehouses.
- Auto Separate is capped at 14 balance weeks.
- Firm PO at Balance Week subtracts cumulative Firm from Firm PO Week through Balance Week.
- Optional HB freeze and FIRM = 0 priority are supported.
- Output includes the WN3-format UPLOAD sheet.
""")


def render_vendor_mapping(psw_files, due_files):
    if not psw_files:
        st.info("Upload PSW / Production Schedule first so the app can detect vendor order for DueDateCalc mapping.")
        return
    payloads = tuple((f.name, bytes(f.getbuffer())) for f in psw_files)
    vendor_rows = detect_vendor_order_cached(payloads)
    if not vendor_rows:
        st.warning("No vendor code was detected from the uploaded PSW files.")
        return
    st.subheader("Detected vendor order for DueDateCalc mapping")
    st.caption("DueDateCalc #1 → Vendor #1, DueDateCalc #2 → Vendor #2, etc.")
    st.dataframe(vendor_rows, hide_index=True, width="stretch")
    due_count = len(due_files or [])
    vendor_count = len(vendor_rows)
    if due_count == 1 and vendor_count > 1:
        st.info("One DueDateCalc is uploaded, so all detected vendors will use the same transit mapping.")
    elif due_count and due_count < vendor_count:
        st.warning(f"{vendor_count} vendors detected but only {due_count} DueDateCalc files are uploaded. The last uploaded DueDateCalc will be used as fallback for remaining vendors.")


st.title("Destination Change App")
st.caption("PlanDetailTimeline + PSW / Production Schedule + DueDateCalc → Optimized output")
render_logic_summary()

left, right = st.columns([1.2, 0.8])
with left:
    st.subheader("1. Upload input files")
    plan_file = st.file_uploader("PlanDetailTimeline raw CSV", type=["csv"], key="plan_file")
    psw_files = st.file_uploader("PSW / Production Schedule.csv files", type=["csv"], accept_multiple_files=True, key="psw_files")
    due_files = st.file_uploader("DueDateCalc.xlsx files", type=["xlsx", "xlsm", "xls"], accept_multiple_files=True, key="due_files")

with right:
    st.subheader("2. Week setup")
    default_current = saturday_of_current_week()
    default_target = default_current + timedelta(days=14)
    target_week = st.date_input("Target Week / Firm PO Week", value=default_target, format="MM/DD/YYYY")
    current_week = st.date_input("Current Week", value=default_current, format="MM/DD/YYYY")
    output_name = st.text_input("Output file name", value=f"destination_change_{date.today().strftime('%Y%m%d')}.xlsx")
    use_osqp = st.checkbox("Add optional OSQP second-pass sheets", value=False)
    auto_balance_week = st.checkbox("Auto separate Firm PO week and balancing week", value=False)
    healthy_ss_pct = st.number_input("Healthy SS% Threshold", min_value=1.0, max_value=1000.0, value=150.0, step=5.0)
    freeze_hold_buy = st.checkbox("Exclude Hold Buy (HB) from destination change", value=False)

render_vendor_mapping(psw_files, due_files)

st.subheader("3. Optional priority rules")
st.caption("Rank is optional. SI = 0 is a hard lock. FIRM = 0 forces Firm PO After = 0 for that warehouse.")
priority_editor = st.data_editor(
    [{"Whse": "", "Mode": "SI", "Value": None, "Rank": None}],
    num_rows="dynamic",
    width="stretch",
    column_config={
        "Whse": st.column_config.TextColumn("Whse"),
        "Mode": st.column_config.SelectboxColumn("Mode", options=["SI", "SS", "FIRM"]),
        "Value": st.column_config.NumberColumn("Value"),
        "Rank": st.column_config.NumberColumn("Rank"),
    },
    key="priority_editor",
)
respect_priority_rank = st.checkbox("Respect Priority Rank", value=False)

if st.button("Run Full Flow", type="primary", width="stretch"):
    missing = []
    if plan_file is None: missing.append("PlanDetailTimeline raw CSV")
    if not psw_files: missing.append("PSW / Production Schedule.csv")
    if not due_files: missing.append("DueDateCalc.xlsx")
    if missing:
        st.error("Missing input: " + ", ".join(missing))
        st.stop()
    if current_week > target_week:
        st.error("Current Week cannot be later than Target Week.")
        st.stop()
    if not output_name.lower().endswith(".xlsx"):
        output_name += ".xlsx"

    import pandas as pd
    from destination_change_unified_flow import PriorityRule, normalize_pct, process_files

    raw_priority = priority_editor
    priority_df = raw_priority if isinstance(raw_priority, pd.DataFrame) else pd.DataFrame(raw_priority)
    rules = {}
    for _, row in priority_df.iterrows():
        whse = normalize_whse(row.get("Whse", ""))
        mode = str(row.get("Mode", "")).strip().upper()
        value = row.get("Value")
        rank = row.get("Rank")
        if not whse or mode not in {"SI", "SS", "FIRM"}:
            continue
        if mode == "FIRM":
            value_float = 0.0
        else:
            if pd.isna(value):
                continue
            value_float = normalize_pct(float(value))
        try:
            rank_int = int(rank) if not pd.isna(rank) else 9999
        except Exception:
            rank_int = 9999
        rules[whse] = PriorityRule(whse=whse, mode=mode, value=value_float, rank=rank_int)

    progress = st.progress(0, text="Preparing files...")
    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            plan_path = save_uploaded(plan_file, tmpdir, "PlanDetailTimeline.csv")
            psw_paths = [save_uploaded(f, tmpdir, f"PSW_{i}.csv") for i, f in enumerate(psw_files, 1)]
            due_paths = [save_uploaded(f, tmpdir, f"DueDateCalc_{i}.xlsx") for i, f in enumerate(due_files, 1)]
            output_path = os.path.join(tmpdir, output_name)
            progress.progress(15, text="Reading vendor / transit mappings...")
            final_path = process_files(
                plan_detail_csv=plan_path,
                production_schedule_csv=psw_paths[0],
                due_date_calc_xlsx=due_paths[0],
                output_path=output_path,
                target_week=target_week,
                current_week=current_week,
                priority_rules=rules,
                psw_csv_paths=psw_paths,
                due_date_calc_xlsx_list=due_paths,
                respect_priority_rank=respect_priority_rank,
                use_osqp_second_pass=use_osqp,
                auto_balance_week=auto_balance_week,
                healthy_ss_pct=healthy_ss_pct,
                freeze_hold_buy=freeze_hold_buy,
            )
            progress.progress(100, text="Completed")
            with open(final_path, "rb") as f:
                st.session_state["last_output_bytes"] = f.read()
            st.session_state["last_output_name"] = Path(final_path).name
            st.success("Destination Change completed successfully.")
        except Exception as exc:
            progress.empty()
            st.error(f"Processing failed: {exc}")
            st.stop()

if "last_output_bytes" in st.session_state:
    st.subheader("Output")
    st.download_button("Download Optimized Excel", data=st.session_state["last_output_bytes"], file_name=st.session_state.get("last_output_name", "destination_change_output.xlsx"), mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
