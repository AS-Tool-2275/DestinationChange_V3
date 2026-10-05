"""
Destination Change - Fast Streamlit UI

The page avoids importing the heavy backend during initial page load.
The optimizer backend is imported only when Run Full Flow is pressed.
"""
from __future__ import annotations

import csv
import io
import os
import tempfile
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Tuple

import streamlit as st

# Must be the first Streamlit command.
st.set_page_config(page_title="Destination Change App", page_icon="📦", layout="wide")


# ------------------------------------------------------------------
# Lightweight UI helpers - no pandas/numpy/openpyxl/backend import here
# ------------------------------------------------------------------


def fmt_date(d: date) -> str:
    return f"{d.month}/{d.day}/{d.year}"


def saturday_of_current_week(today: date | None = None) -> date:
    if today is None:
        today = date.today()
    return today + timedelta(days=(5 - today.weekday()) % 7)


def normalize_whse(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return ""
    try:
        f = float(text)
        if f.is_integer():
            return str(int(f))
    except Exception:
        pass
    return text.upper()


def normalize_pct(value) -> float:
    value = float(value)
    return value / 100.0 if value > 1 or value < -1 else value


def _backend():
    # Deliberately lazy: do not load pandas/numpy/openpyxl/optimizer on page startup.
    from destination_change_unified_flow import process_files
    return process_files


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


def _normalize_vendor(value) -> str:
    if value is None:
        return ""
    text = str(value).strip().strip('"').upper()
    if not text or text.lower() == "nan":
        return ""
    # Keep simple numeric vendor codes stable.
    try:
        f = float(text)
        if f.is_integer():
            return str(int(f))
    except Exception:
        pass
    return text


def _vendor_match_key(value) -> str:
    import re
    text = _normalize_vendor(value)
    if not text:
        return ""
    m = re.search(r"\((0*\d+)\)", text)
    if m:
        return str(int(m.group(1)))
    nums = re.findall(r"0*\d+", text)
    if nums:
        return str(int(nums[-1]))
    return text


def detect_psw_vendors_fast(psw_payload: Tuple[Tuple[str, bytes], ...]) -> List[Dict[str, object]]:
    """Lightweight vendor detection using csv only; backend is not imported."""
    rows_out: List[Dict[str, object]] = []
    seen = set()
    order = 1
    for file_order, (name, payload) in enumerate(psw_payload, start=1):
        text = payload.decode("utf-8-sig", errors="replace")
        reader = csv.reader(io.StringIO(text))
        header = None
        header_idx = None
        for i, row in enumerate(reader):
            if row and str(row[0]).strip().lower() == "item #":
                header = [str(x).strip() for x in row]
                header_idx = i
                break
        if header is None:
            continue
        vendor_idx = None
        for i, col in enumerate(header):
            low = col.lower()
            if low in {"vendor", "vendor code", "vendorcode", "vendor #", "vendor#", "supplier", "supplier code"} or "vendor" in low or "supplier" in low:
                vendor_idx = i
                break
        if vendor_idx is None:
            continue

        row_counts: Dict[str, int] = {}
        # Resume from the row after header without storing the file in a DataFrame.
        for row in reader:
            if len(row) <= vendor_idx:
                continue
            vendor = _normalize_vendor(row[vendor_idx])
            key = _vendor_match_key(vendor)
            if not key:
                continue
            row_counts[key] = row_counts.get(key, 0) + 1
            if key not in seen:
                seen.add(key)
                rows_out.append({
                    "Vendor Order": order,
                    "Vendor Code": key,
                    "Source PSW File Order": file_order,
                    "Rows": 0,
                })
                order += 1
        for item in rows_out:
            if item["Source PSW File Order"] == file_order:
                item["Rows"] = row_counts.get(item["Vendor Code"], item["Rows"])
    return rows_out


def build_priority_rules(priority_data):
    """Convert Streamlit editor output to backend PriorityRule objects."""
    import math
    from destination_change_unified_flow import PriorityRule
    rules: Dict[str, PriorityRule] = {}
    if priority_data is None:
        return rules

    try:
        records = priority_data.to_dict("records")
    except AttributeError:
        records = list(priority_data)

    for row in records:
        whse = normalize_whse(row.get("Whse", ""))
        mode = str(row.get("Mode", "")).strip().upper()
        value = row.get("Value")
        rank = row.get("Rank")
        if not whse or mode not in {"SI", "SS", "FIRM"}:
            continue
        if mode == "FIRM":
            value_float = 0.0
        else:
            try:
                if value is None or (isinstance(value, float) and math.isnan(value)):
                    continue
                value_float = normalize_pct(float(value))
            except Exception:
                continue
        try:
            rank_int = int(rank) if rank is not None and not (isinstance(rank, float) and math.isnan(rank)) else 9999
        except Exception:
            rank_int = 9999
        rules[whse] = PriorityRule(whse=whse, mode=mode, value=value_float, rank=rank_int)
    return rules


def render_logic_summary() -> None:
    with st.expander("Logic summary", expanded=False):
        st.markdown(
            """
**Inputs**
- PlanDetailTimeline.csv: inventory planning / ETA timeline.
- PSW / Production Schedule.csv: firm supply source. Only S/F/P = F is used for Firm PO.
- DueDateCalc.xlsx: warehouse delivery days converted to whole-week offsets with CEILING(days / 7).

**Optimization**
- Main-vendor Firm PO is optimized first.
- Optional Hold Buy (HB) freeze.
- Priority SI = 0 hard lock; FIRM = 0 forces Firm PO After = 0.
- Auto Separate uses all eligible Buy warehouses and caps Balance Week at 14 weeks.
- Firm PO Week is user-selected; SI at Balance Week subtracts all Firm from Firm PO Week through Balance Week, inclusive.
- Multi-vendor / sub-vendor flow and 9 PlanDetailTimeline metadata columns are preserved.
- Workbook includes WN3-format UPLOAD output for non-zero destination changes, including zero-quantity donor rows.
"""
        )


def render_vendor_mapping(psw_files, due_files) -> None:
    if not psw_files:
        st.info("Upload PSW / Production Schedule first so the app can detect vendor order for DueDateCalc mapping.")
        return

    try:
        # Recompute only when PSW file names/sizes change.
        signature = tuple((f.name, getattr(f, "size", len(f.getbuffer()))) for f in psw_files)
        cached_sig = st.session_state.get("vendor_signature")
        if cached_sig != signature:
            payload = tuple((f.name, bytes(f.getbuffer())) for f in psw_files)
            st.session_state["vendor_rows"] = detect_psw_vendors_fast(payload)
            st.session_state["vendor_signature"] = signature
        vendor_rows = st.session_state.get("vendor_rows", [])

        if not vendor_rows:
            st.warning("No vendor code was detected from the uploaded PSW files.")
            return

        st.subheader("Detected vendor order for DueDateCalc mapping")
        st.caption("Upload DueDateCalc files in this vendor order: DueDateCalc #1 → Vendor #1, DueDateCalc #2 → Vendor #2, etc.")
        st.dataframe(vendor_rows, width="stretch", hide_index=True)

        due_count = len(due_files or [])
        vendor_count = len(vendor_rows)
        if due_count == 1 and vendor_count > 1:
            st.info("One DueDateCalc is uploaded, so all detected vendors will use the same transit mapping.")
        elif due_count and due_count < vendor_count:
            st.warning(f"{vendor_count} vendors detected but only {due_count} DueDateCalc files are uploaded. The last uploaded DueDateCalc will be used as fallback.")
    except Exception as exc:
        st.warning(f"Could not detect vendor order yet: {exc}")


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
    output_name = st.text_input("Output file name", value=f"destination_change_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    use_osqp = st.checkbox("Add optional OSQP second-pass sheets", value=False)
    auto_balance_week = st.checkbox("Auto separate Firm PO week and balancing week", value=False)
    healthy_ss_pct = st.number_input("Healthy SS% Threshold", min_value=1.0, max_value=1000.0, value=150.0, step=5.0)
    freeze_hold_buy = st.checkbox("Exclude Hold Buy (HB) from destination change", value=False)

render_vendor_mapping(psw_files, due_files)

st.subheader("3. Optional priority rules")
st.caption("Rank is optional. SI = 0 is a hard lock. FIRM = 0 forces Firm PO After = 0 for that warehouse.")
priority_df = st.data_editor(
    [{"Whse": "", "Mode": "SI", "Value": None, "Rank": None}],
    num_rows="dynamic",
    width="stretch",
    column_config={
        "Whse": st.column_config.TextColumn("Whse"),
        "Mode": st.column_config.SelectboxColumn("Mode", options=["SI", "SS", "FIRM"]),
        "Value": st.column_config.NumberColumn("Value", help="SI/SS: 50 or 0.5 = 50%; FIRM: 0"),
        "Rank": st.column_config.NumberColumn("Rank", help="1 = highest priority"),
    },
    key="priority_editor",
)
respect_priority_rank = st.checkbox("Respect Priority Rank", value=False)

if st.button("Run Full Flow", type="primary", width="stretch"):
    missing = []
    if plan_file is None:
        missing.append("PlanDetailTimeline raw CSV")
    if not psw_files:
        missing.append("PSW / Production Schedule.csv")
    if not due_files:
        missing.append("DueDateCalc.xlsx")
    if missing:
        st.error("Missing input: " + ", ".join(missing))
        st.stop()
    if current_week > target_week:
        st.error("Current Week cannot be later than Target Week.")
        st.stop()
    if not output_name.lower().endswith(".xlsx"):
        output_name += ".xlsx"

    priority_rules = build_priority_rules(priority_df)
    progress = st.progress(0, text="Preparing files...")
    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            plan_path = save_uploaded(plan_file, tmpdir, "PlanDetailTimeline.csv")
            psw_paths = [save_uploaded(f, tmpdir, f"PSW_{i}.csv") for i, f in enumerate(psw_files, 1)]
            due_paths = [save_uploaded(f, tmpdir, f"DueDateCalc_{i}.xlsx") for i, f in enumerate(due_files, 1)]
            output_path = os.path.join(tmpdir, output_name)

            progress.progress(20, text="Loading optimizer backend...")
            process_files = _backend()
            progress.progress(35, text="Calculating inventory and Firm PO...")
            final_path = process_files(
                plan_detail_csv=plan_path,
                production_schedule_csv=psw_paths[0],
                due_date_calc_xlsx=due_paths[0],
                output_path=output_path,
                target_week=target_week,
                current_week=current_week,
                priority_rules=priority_rules,
                psw_csv_paths=psw_paths,
                due_date_calc_xlsx_list=due_paths,
                respect_priority_rank=respect_priority_rank,
                use_osqp_second_pass=use_osqp,
                auto_balance_week=auto_balance_week,
                healthy_ss_pct=healthy_ss_pct,
                freeze_hold_buy=freeze_hold_buy,
            )
            progress.progress(90, text="Preparing download...")
            with open(final_path, "rb") as f:
                output_bytes = f.read()
            st.session_state["last_output_bytes"] = output_bytes
            st.session_state["last_output_name"] = Path(final_path).name
            st.session_state["last_run_info"] = {
                "Firm PO Week": fmt_date(target_week),
                "Current Week": fmt_date(current_week),
                "Priority Rules": len(priority_rules),
                "Priority Rank Enabled": respect_priority_rank,
                "Auto Balancing Week": "Enabled" if auto_balance_week else "Disabled",
                "OSQP Second Pass": use_osqp,
                "Hold Buy Freeze": freeze_hold_buy,
            }
            progress.progress(100, text="Completed")
            st.success("Destination Change completed successfully.")
        except Exception as exc:
            progress.empty()
            st.error(f"Processing failed: {exc}")
            st.stop()

if "last_output_bytes" in st.session_state:
    st.subheader("Output")
    st.download_button("Download Optimized Excel", data=st.session_state["last_output_bytes"], file_name=st.session_state.get("last_output_name", "destination_change_output.xlsx"), mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
    with st.expander("Run information", expanded=False):
        st.json(st.session_state.get("last_run_info", {}))
