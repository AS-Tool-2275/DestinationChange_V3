"""
Destination Change Unified Flow - Streamlit Application

Main entry point for Streamlit Cloud or an internal Streamlit server.
The backend module must be stored in the same folder.
"""

from __future__ import annotations

import os
import tempfile
from datetime import date, timedelta
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st


def fmt_date(d: date) -> str:
    return f"{d.month}/{d.day}/{d.year}"


def saturday_of_current_week(today: date | None = None) -> date:
    if today is None:
        today = date.today()
    return today + timedelta(days=(5 - today.weekday()) % 7)


def normalize_whse(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
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
    # Lazy import: avoid loading numpy/openpyxl/optimizer code on the initial Streamlit page load.
    from destination_change_unified_flow import (
        PriorityRule,
        detect_psw_vendors,
        process_files,
    )
    return PriorityRule, detect_psw_vendors, process_files

st.set_page_config(
    page_title="Destination Change App",
    page_icon="📦",
    layout="wide",
)


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


def build_priority_rules(priority_df: pd.DataFrame):
    PriorityRule, _, _ = _backend()
    rules: Dict[str, PriorityRule] = {}
    if priority_df is None or priority_df.empty:
        return rules

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
            try:
                value_float = normalize_pct(float(value))
            except Exception:
                continue
        try:
            rank_int = int(rank) if not pd.isna(rank) else 9999
        except Exception:
            rank_int = 9999
        rules[whse] = PriorityRule(
            whse=whse,
            mode=mode,
            value=value_float,
            rank=rank_int,
        )
    return rules


def render_logic_summary() -> None:
    with st.expander("Logic summary", expanded=False):
        st.markdown(
            """
**Inputs**
- PlanDetailTimeline.csv: inventory planning / ETA timeline.
- PSW / Production Schedule.csv: firm supply source. Only S/F/P = F is used for Firm PO.
- DueDateCalc.xlsx: warehouse delivery days converted to whole-week offsets with CEILING(days / 7).

**Vendor-aware DueDateCalc mapping**
- Vendor order is detected from PSW / Production Schedule.
- DueDateCalc #1 maps to Vendor #1, DueDateCalc #2 to Vendor #2, and so on.
- If only one DueDateCalc is uploaded, all detected vendors use the same transit file.
- If fewer DueDateCalc files than vendors are uploaded, the last uploaded file is used as fallback.

**Optimization**
- Main-vendor Firm PO is optimized first.
- Priority rules are optional. Priority Rank is optional and can be enabled independently.
- Priority SI = 0 is a hard lock: original Firm PO quantity is preserved for that warehouse.
- Remaining allocation uses Lowest SS% After -> Highest SI After -> Warehouse code.
- Zero-SS items receive an automatic fallback SI equalization pass after the main allocation.
- Sub-vendor suggestion mirrors the main-vendor allocation logic, but starts from Main Vendor SI After.
- The output preserves total Firm PO quantity by item.

**Optional automatic balancing week**
- Firm PO week is the selected Target Week. Candidate Balance Weeks are checked sequentially through the last fully evaluable PlanDetailTimeline horizon; there is no fixed Week-14 cap.
- Reference warehouse / Adjusted Healthy logic is removed. Healthy SS% is always the fixed user threshold.
- Every candidate week is rebuilt fresh from PlanDetailTimeline and Destination Change is rerun across all Buy (B) warehouses.
- **Healthy SS%** is the normal-planning gate: Active WH = driver-vendor Firm After > 0, and every Active WH with meaningful SS must be at/below Healthy SS%.
- **Future shortage protection** scans the remaining evaluable horizon for all change-eligible Buy warehouses. Upcoming demand peaks create a minimum Future Firm Floor before normal SS balancing; a candidate is rejected if projected SI still drops below 0.
- **Runout balance** handles phase-out items: the first zero-SS bucket uses Last Positive SS as a bridge; from 2 consecutive zero-SS buckets onward, SS% no longer gates the candidate and projected future SI >= 0 is used instead.
- Firm After = 0 is excluded only from the Healthy SS% gate for that candidate. It may become Active again next week and can still be protected by the future-shortage guard when change-eligible.
- Main Vendor drives when Main Firm > 0. If Main Firm = 0 and Sub Firm > 0, Sub Vendor uses the same Hybrid logic with Main Vendor SI After as its baseline.
- Firm PO Week remains the selected Target Week; SI at Balance Week subtracts all Firm PO from Firm PO Week through Balance Week, inclusive.
- Warehouse 335 uses accumulated NET FCST from Current Week through the selected balancing week without ETA->ETD offset shifting.
- Priority / Hard Lock / FIRM=0 / optional Hold Buy (HB) freeze rules continue to apply inside each candidate allocation.
- The downloaded workbook includes an UPLOAD sheet in WN3 format: main-vendor rows first, then sub-vendor rows; only non-zero destination-change rows are included.
"""
        )


@st.cache_data(show_spinner=False, max_entries=8)
def detect_psw_vendors_cached(files_payload: Tuple[Tuple[str, bytes], ...]) -> pd.DataFrame:
    _, detect_psw_vendors, _ = _backend()
    with tempfile.TemporaryDirectory() as tmpdir:
        paths = []
        for i, (name, payload) in enumerate(files_payload, 1):
            safe_name = Path(name or f"PSW_{i}.csv").name
            path = Path(tmpdir) / f"{i}_{safe_name}"
            path.write_bytes(payload)
            paths.append(str(path))
        return detect_psw_vendors(paths)


def render_vendor_mapping(psw_files) -> None:
    if not psw_files:
        st.info("Upload PSW / Production Schedule first so the app can detect vendor order for DueDateCalc mapping.")
        return

    try:
        vendor_df = detect_psw_vendors_cached(tuple((f.name, bytes(f.getbuffer())) for f in psw_files))
        if vendor_df.empty:
            st.warning("No vendor code was detected from the uploaded PSW files.")
            return

        display_df = vendor_df[["Vendor Order", "Vendor Code", "Source PSW File Order", "Rows"]].copy()
        st.subheader("Detected vendor order for DueDateCalc mapping")
        st.caption(
            "Upload DueDateCalc files in this vendor order: DueDateCalc #1 → Vendor #1, "
            "DueDateCalc #2 → Vendor #2, etc."
        )
        st.dataframe(display_df, width="stretch", hide_index=True)

        due_count = len(due_files or [])
        vendor_count = len(vendor_df)
        if due_count == 1 and vendor_count > 1:
            st.info("One DueDateCalc is uploaded, so all detected vendors will use the same transit mapping.")
        elif due_count and due_count < vendor_count:
            st.warning(
                f"{vendor_count} vendors detected but only {due_count} DueDateCalc files are uploaded. "
                "The last uploaded DueDateCalc will be used as fallback for the remaining vendors."
            )
    except Exception as exc:
        st.warning(f"Could not detect vendor order yet: {exc}")


st.title("Destination Change App")
st.caption("PlanDetailTimeline + PSW / Production Schedule + DueDateCalc → Optimized output")
render_logic_summary()

left, right = st.columns([1.2, 0.8])

with left:
    st.subheader("1. Upload input files")
    plan_file = st.file_uploader("PlanDetailTimeline raw CSV", type=["csv"], key="plan_file")
    psw_files = st.file_uploader(
        "PSW / Production Schedule.csv files",
        type=["csv"],
        accept_multiple_files=True,
        key="psw_files",
        help="Upload one or more PSW files. Vendor order is detected from these files.",
    )
    due_files = st.file_uploader(
        "DueDateCalc.xlsx files",
        type=["xlsx", "xlsm", "xls"],
        accept_multiple_files=True,
        key="due_files",
        help="DueDateCalc upload order must follow the detected vendor order from PSW.",
    )

with right:
    st.subheader("2. Week setup")
    default_current = saturday_of_current_week()
    default_target = default_current + timedelta(days=14)
    target_week = st.date_input("Target Week / Firm PO Week", value=default_target, format="MM/DD/YYYY")
    current_week = st.date_input("Current Week", value=default_current, format="MM/DD/YYYY")
    output_name = st.text_input(
        "Output file name",
        value=f"destination_change_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    )
    use_osqp = st.checkbox(
        "Add optional OSQP second-pass sheets",
        value=False,
        help="Adds separate main-vendor and sub-vendor OSQP what-if sheets without replacing Optimized Data.",
    )
    auto_balance_week = st.checkbox(
        "Auto separate Firm PO week and balancing week",
        value=False,
        help=(
            "Off: Firm PO Week is also the balancing week. "
            "On: each candidate week is rebuilt fresh across all Buy warehouses. "
            "Normal planning uses fixed Healthy SS%; future shortage protection catches upcoming demand peaks; "
            "and phase-out items switch to Runout Balance when SS stays at 0. Firm After = 0 is excluded only from the Healthy SS% gate for that candidate. "
            "The search continues through the last fully evaluable PlanDetailTimeline horizon instead of a fixed Week-14 cap."
        ),
    )
    healthy_ss_pct = st.number_input(
        "Healthy SS% Threshold",
        min_value=1.0,
        max_value=1000.0,
        value=150.0,
        step=5.0,
        help="Used only by Auto separate Firm PO week and balancing week. Default = 150%.",
    )
    freeze_hold_buy = st.checkbox(
        "Exclude Hold Buy (HB) from destination change",
        value=False,
        help="When enabled, Item + Warehouse rows with Hold/ Buy = HB keep original Firm PO and are excluded from destination change.",
    )

render_vendor_mapping(psw_files)

st.subheader("3. Optional priority rules")
st.caption(
    "Rank is optional. When enabled, the lowest Rank is considered first; warehouses with the same Rank are considered together. "
    "SI = 0 is a hard lock. FIRM = 0 forces Firm PO After = 0 for that warehouse."
)

priority_df = st.data_editor(
    pd.DataFrame(columns=["Whse", "Mode", "Value", "Rank"]),
    num_rows="dynamic",
    width="stretch",
    column_config={
        "Whse": st.column_config.TextColumn("Whse"),
        "Mode": st.column_config.SelectboxColumn("Mode", options=["SI", "SS", "FIRM"]),
        "Value": st.column_config.NumberColumn("Value", help="SI/SS: 50 or 0.5 = 50%; FIRM: use 0 to force Firm PO After = 0"),
        "Rank": st.column_config.NumberColumn("Rank", help="1 = highest priority; same rank = same priority"),
    },
    key="priority_editor",
)

respect_priority_rank = st.checkbox(
    "Respect Priority Rank",
    value=False,
    help="Off = current priority metric logic. On = lower Rank is processed first; same Rank is treated as the same priority group.",
)

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

            progress.progress(15, text="Reading vendor / transit mappings...")
            progress.progress(35, text="Calculating inventory and Firm PO...")
            _, _, process_files = _backend()
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
    st.download_button(
        "Download Optimized Excel",
        data=st.session_state["last_output_bytes"],
        file_name=st.session_state.get("last_output_name", "destination_change_output.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )
    with st.expander("Run information", expanded=False):
        st.json(st.session_state.get("last_run_info", {}))
