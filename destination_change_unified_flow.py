"""
Destination Change Unified Flow

Input:
  1) PlanDetailTimeline raw CSV (including the first 6 metadata rows)
  2) Production Schedule raw CSV (including the first 6 metadata rows)
  3) DueDateCalc Excel
  4) Target Week / Wk3

Output:
  Final optimized Excel with debug sheets for logic review.

Default behavior:
  - F Wk3 is taken from Production Schedule / PSW, using only S/F/P = F for the main vendor.
  - PlanDetailTimeline is converted from ETA to ETD using DueDateCalc.
  - Warehouse offsets are always calculated directly from DueDateCalc using ceil(Delivery Days / 7).
"""

from __future__ import annotations

import argparse
import math
import os
import re
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd
import numpy as np
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, simpledialog
except Exception:
    tk = None
    ttk = None
    filedialog = None
    messagebox = None
    simpledialog = None

# ============================================================
# Config
# ============================================================

OUTPUT_COLUMNS = [
    "Item",
    "ProdResourceID",
    "Whse",
    "F Wk3",
    "Sum of SI Wk3",
    "Sum of SI-SS Wk3",
    "Average of SS Wk3",
    "Vendor",
    # Multi-vendor audit columns. These are optional and are populated when PSW vendor detail is available.
    "Main Vendor",
    "Main Vendor F Wk3",
    "Other Vendor Supply",
    "Other Vendor List",
    "Timeline Firm PO",
    "PSW F Used for Reconciliation",
    "Firm PO Reconciliation Gap",
    "Total Supply Added to SI",
 ]

PLAN_METADATA_COLUMNS = [
    "Item Class",
    "Coll. Class",
    "Series",
    "Division",
    "Item Status",
    "Future Status",
    "Hold/ Buy",
    "ABC",
    "Source Key",
]

DTYPE_MAP = {
    "FIRM DEMAND": "FIRM DEMANDS",
    "FIRM DEMANDS": "FIRM DEMANDS",
    "FIRM POS": "FIRM POS",
    "FIRM PO": "FIRM POS",
    "PLANNED POS": "PLANNED POS",
    "PLANNED PO": "PLANNED POS",
    "SHIPPABLE INV": "SHIPPABLE INV",
    "SHIPPABLE INVENTORY": "SHIPPABLE INV",
    "SAFETY STK": "SAFETY STK",
    "SAFETY STOCK": "SAFETY STK",
    "NET FCST": "NET FCST",
    "NET FORECAST": "NET FCST",
}

# Mapping currently used by SI-SS_WANEK 3.py.
# This keeps the output aligned with the current logic when the current DueDateCalc is used.

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
VENDOR_FALLBACK_COL = "Vendor"

# Auto Separate uses the available PlanDetailTimeline horizon dynamically.
# There is no fixed Week-14 cap. Candidate weeks continue through the last
# evaluable PlanDetail week for all mapped Buy warehouses.


# ============================================================
# Common helpers
# ============================================================

def normalize_item(value) -> str:
    """Clean Item # values like ="01226" -> 1226."""
    if pd.isna(value):
        return ""
    text = str(value).strip()
    m = re.match(r'^=\s*"(.*)"$', text)
    if m:
        text = m.group(1).strip()
    text = text.strip().strip('"').strip()
    # Ashley exports often keep leading zero as Excel formula text. Existing output uses integer-like item.
    if re.fullmatch(r"0*\d+", text):
        return str(int(text))
    return text


def normalize_whse(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    m = re.match(r'^=\s*"(.*)"$', text)
    if m:
        text = m.group(1).strip()
    try:
        f = float(text)
        if f.is_integer():
            return str(int(f))
    except Exception:
        pass
    return text.strip().upper()


def clean_dtype(series: pd.Series) -> pd.Series:
    s = series.fillna("").astype(str).str.strip().str.upper()
    return s.map(lambda x: DTYPE_MAP.get(x, x))


def parse_user_date(text: str) -> date:
    text = str(text).strip()
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return pd.to_datetime(text).date()


def fmt_date(d: date) -> str:
    return f"{d.month}/{d.day}/{d.year}"


def saturday_of_current_week(today: Optional[date] = None) -> date:
    if today is None:
        today = date.today()
    return today + timedelta(days=(5 - today.weekday()) % 7)


def ensure_unique_output_path(path: str) -> str:
    path = os.path.abspath(path)
    if not os.path.exists(path):
        return path
    folder, filename = os.path.split(path)
    stem, ext = os.path.splitext(filename)
    idx = 1
    while True:
        candidate = os.path.join(folder, f"{stem}_{idx}{ext}")
        if not os.path.exists(candidate):
            return candidate
        idx += 1


def read_report_csv(path: str, dtype=str) -> pd.DataFrame:
    """Read Ashley report CSV that has metadata lines before actual header."""
    header_row = None
    with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
        for i, line in enumerate(f):
            if line.lstrip().startswith("Item #"):
                header_row = i
                break
    if header_row is None:
        raise ValueError(f"Could not find the header row starting with 'Item #' in file: {path}")
    df = pd.read_csv(path, skiprows=header_row, dtype=dtype, low_memory=False, index_col=False)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def extract_report_date_from_csv(path: str) -> Optional[datetime]:
    with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
        for _ in range(10):
            line = f.readline()
            if not line:
                break
            if "Report Date" in line:
                text = line.split(":", 1)[-1].strip()
                try:
                    return pd.to_datetime(text).to_pydatetime()
                except Exception:
                    return None
    return None


def parse_header_to_date(col_name) -> Optional[date]:
    if isinstance(col_name, (datetime, pd.Timestamp)):
        return pd.to_datetime(col_name).date()
    text = str(col_name).strip()
    try:
        return pd.to_datetime(text).date()
    except Exception:
        return None


def build_date_column_map(df: pd.DataFrame) -> Dict[date, str]:
    mapping = {}
    for c in df.columns:
        d = parse_header_to_date(c)
        if d is not None:
            mapping[d] = c
    return mapping


def date_range_saturdays(start_date: date, end_date: date) -> List[date]:
    out = []
    cur = start_date
    while cur <= end_date:
        out.append(cur)
        cur += timedelta(days=7)
    return out


def get_numeric(df: pd.DataFrame, col) -> pd.Series:
    if col not in df.columns:
        return pd.Series(0.0, index=df.index)
    return pd.to_numeric(df[col], errors="coerce").fillna(0.0)


def group_value(df: pd.DataFrame, key_cols, value_col, output_name) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=key_cols + [output_name])
    out = df.groupby(key_cols, dropna=False, as_index=False)[value_col].sum()
    return out.rename(columns={value_col: output_name})


# ============================================================
# Step 1: DueDateCalc -> warehouse offset
# ============================================================

def load_due_date_offsets(due_date_calc_path: str) -> Tuple[Dict[str, int], pd.DataFrame]:
    """Read warehouse delivery days and convert them to whole-week offsets.

    The application uses DueDateCalc directly. Offset Weeks = CEILING(Delivery Days / 7).
    """
    raw = pd.read_excel(due_date_calc_path, sheet_name=0, header=None)
    header_idx = None
    for i in range(len(raw)):
        vals = [str(x).strip() for x in raw.iloc[i].tolist()]
        if "Warehouse" in vals and any("Delivery Days" in v for v in vals):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("Could not find Warehouse / Delivery Days header in DueDateCalc.")

    df = pd.read_excel(due_date_calc_path, sheet_name=0, header=header_idx)
    df.columns = [str(c).strip() for c in df.columns]
    delivery_col = next((c for c in df.columns if "Delivery Days" in c), None)
    if delivery_col is None:
        raise ValueError("DueDateCalc is missing the Delivery Days column.")

    rows = []
    offset_map: Dict[str, int] = {}
    for _, r in df.iterrows():
        warehouse_text = str(r.get("Warehouse", "")).strip()
        if not warehouse_text or warehouse_text.lower() == "nan":
            continue
        whse = normalize_whse(warehouse_text.split("-", 1)[0])
        if not whse:
            continue
        days = pd.to_numeric(r.get(delivery_col), errors="coerce")
        if pd.isna(days):
            continue
        used_offset = max(1, int(math.ceil(float(days) / 7.0)))
        offset_map[whse] = used_offset
        rows.append({
            "Whse": whse,
            "Warehouse": warehouse_text,
            "Delivery Days": float(days),
            "Used Offset Weeks": used_offset,
            "Offset Source": "DueDateCalc ceil(days/7)",
        })
    if not offset_map:
        raise ValueError("Could not read warehouse offsets from DueDateCalc.")
    return offset_map, pd.DataFrame(rows)


# ============================================================
# Step 2: PlanDetailTimeline ETA -> ETD
# ============================================================

def convert_plan_eta_to_etd(plan_csv_path: str, offset_map: Dict[str, int]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    raw = read_report_csv(plan_csv_path, dtype=str)
    required = ["Item #", "Whse", "Data Type"]
    missing = [c for c in required if c not in raw.columns]
    if missing:
        raise ValueError(f"PlanDetailTimeline is missing required columns: {missing}")

    raw = raw.copy()
    raw["Item #"] = raw["Item #"].map(normalize_item)
    raw["Whse"] = raw["Whse"].map(normalize_whse)
    raw["Data Type"] = raw["Data Type"].fillna("").astype(str).str.strip()

    # Robustly detect weekly date columns by header value.
    # Older logic used raw.columns[3:-20], but some PlanDetailTimeline exports have a
    # different number of master-data columns, which can accidentally include fields
    # like "Item Class" in the date range.
    original_date_cols = []
    original_dates = []
    for c in raw.columns:
        d = parse_header_to_date(c)
        if d is not None:
            original_date_cols.append(c)
            original_dates.append(d)
    if not original_date_cols:
        raise ValueError("PlanDetailTimeline does not contain recognizable weekly date columns.")

    # Match SI-SS_WANEK 3.py: extend 22 weeks (154 days) backward.
    extended_dates = [d - timedelta(days=154) for d in original_dates]
    all_dates = extended_dates + original_dates
    date_labels = [fmt_date(d) for d in all_dates]

    out_values = pd.DataFrame(0.0, index=raw.index, columns=date_labels)
    numeric_original = raw[original_date_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)

    unknown_whse = set()
    offset_used_by_row = []
    for whse, idx in raw.groupby("Whse", sort=False).groups.items():
        offset = offset_map.get(whse)
        if offset is None:
            # Safe fallback: no shift. Debug will expose this.
            offset = 0
            unknown_whse.add(whse)
        offset_used_by_row.extend([(i, offset) for i in idx])
        shifted_labels = [fmt_date(d - timedelta(days=7 * offset)) for d in original_dates]
        # Some shifted labels may be outside all_dates if offset > 22; ignore those safely.
        for src_col, dst_label in zip(original_date_cols, shifted_labels):
            if dst_label in out_values.columns:
                out_values.loc[idx, dst_label] = numeric_original.loc[idx, src_col].values

    master_cols = list(raw.columns[-20:]) if len(raw.columns) >= 23 else []
    converted = pd.concat([raw[required].reset_index(drop=True), out_values.reset_index(drop=True), raw[master_cols].reset_index(drop=True)], axis=1)

    debug = pd.DataFrame([
        ["Plan rows", len(raw)],
        ["Original first week", fmt_date(min(original_dates))],
        ["Original last week", fmt_date(max(original_dates))],
        ["Converted first ETD week", fmt_date(min(all_dates))],
        ["Converted last ETD week", fmt_date(max(all_dates))],
        ["Unknown Whse count", len(unknown_whse)],
        ["Unknown Whse list", ", ".join(sorted(unknown_whse))],
    ], columns=["Field", "Value"])
    return converted, debug


# ============================================================
# Step 3: Production Schedule -> F Wk3
# ============================================================

def build_production_date_map(columns: Iterable[str], report_date: Optional[datetime], target_week: date) -> Dict[date, str]:
    date_cols = []
    for c in columns:
        text = str(c).strip()
        if re.fullmatch(r"\d{1,2}/\d{1,2}", text):
            date_cols.append(c)
    if not date_cols:
        raise ValueError("Production Schedule does not contain weekly columns in M/D format.")

    base_year = report_date.year if report_date is not None else target_week.year
    # If the first schedule month is much later than target month, it may belong to previous year.
    first_month = int(str(date_cols[0]).strip().split("/")[0])
    year = base_year
    if first_month - target_week.month > 6:
        year -= 1

    mapping = {}
    prev_month = None
    for c in date_cols:
        m, d = [int(x) for x in str(c).strip().split("/")]
        if prev_month is not None and m < prev_month:
            year += 1
        prev_month = m
        mapping[date(year, m, d)] = c
    return mapping



def find_vendor_col(df: pd.DataFrame) -> Optional[str]:
    """Find the most likely vendor column in PlanDetailTimeline / PSW exports."""
    candidates_exact = [
        "Vendor", "Vendor Code", "VendorCode", "Vendor #", "Vendor#",
        "Vend", "Vend Code", "Supplier", "Supplier Code", "Mfg Vendor", "MFG Vendor",
    ]
    lower_map = {str(c).strip().lower(): c for c in df.columns}
    for name in candidates_exact:
        if name.lower() in lower_map:
            return lower_map[name.lower()]
    for c in df.columns:
        text = str(c).strip().lower()
        if "vendor" in text or "supplier" in text:
            return c
    return None


def normalize_vendor(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    m = re.match(r'^=\s*"(.*)"$', text)
    if m:
        text = m.group(1).strip()
    text = text.strip().strip('"').strip().upper()
    if re.fullmatch(r"0*\d+", text):
        return str(int(text))
    return text




def vendor_match_key(value) -> str:
    """Return comparable vendor code. Timeline may show name(code), PSW may show code only."""
    text = normalize_vendor(value)
    if not text:
        return ""
    paren = re.search(r"\((0*\d+)\)", text)
    if paren:
        return str(int(paren.group(1)))
    nums = re.findall(r"0*\d+", text)
    if nums:
        return str(int(nums[-1]))
    return text


def detect_timeline_vendors(plan_csv_path: str) -> pd.DataFrame:
    """Detect unique vendor order from PlanDetailTimeline for DueDateCalc upload mapping.

    The app uses this list to tell users which DueDateCalc file should correspond
    to each vendor. Vendor order follows first appearance in the Timeline file.
    """
    raw = read_report_csv(plan_csv_path, dtype=str)
    raw.columns = [str(c).strip() for c in raw.columns]
    vendor_col = find_vendor_col(raw)
    if vendor_col is None:
        return pd.DataFrame(columns=["Vendor Order", "Vendor", "Vendor Code", "Vendor Key", "Rows"])

    tmp = raw.copy()
    tmp["_vendor"] = tmp[vendor_col].map(normalize_vendor)
    tmp["_vendor_key"] = tmp["_vendor"].map(vendor_match_key)
    tmp = tmp[(tmp["_vendor"] != "") & (tmp["_vendor_key"] != "")].copy()
    if tmp.empty:
        return pd.DataFrame(columns=["Vendor Order", "Vendor", "Vendor Code", "Vendor Key", "Rows"])

    rows = []
    seen = set()
    order = 1
    for _, r in tmp.iterrows():
        key = str(r["_vendor_key"]).strip()
        if key in seen:
            continue
        seen.add(key)
        vendor = str(r["_vendor"]).strip()
        rows.append({
            "Vendor Order": order,
            "Vendor": vendor,
            "Vendor Code": key,
            "Vendor Key": key,
            "Rows": int((tmp["_vendor_key"] == key).sum()),
        })
        order += 1
    return pd.DataFrame(rows)


def detect_psw_vendors(psw_csv_paths: List[str]) -> pd.DataFrame:
    """Detect unique vendor order from PSW / Production Schedule files.

    This is used by the Streamlit UI and backend to map DueDateCalc upload order
    to vendor-specific transit. Vendor order follows first appearance across the
    uploaded PSW files. If one PSW file contains multiple vendors, each vendor is
    listed separately.
    """
    rows = []
    seen = set()
    order = 1
    for file_order, path in enumerate(psw_csv_paths or [], start=1):
        if not path:
            continue
        try:
            raw = read_report_csv(path, dtype=str)
        except Exception:
            continue
        raw.columns = [str(c).strip() for c in raw.columns]
        vendor_col = find_vendor_col(raw)
        if vendor_col is None:
            continue
        tmp = raw.copy()
        tmp["_vendor"] = tmp[vendor_col].map(normalize_vendor)
        tmp["_vendor_key"] = tmp["_vendor"].map(vendor_match_key)
        tmp = tmp[(tmp["_vendor"] != "") & (tmp["_vendor_key"] != "")].copy()
        for _, r in tmp.iterrows():
            key = str(r["_vendor_key"]).strip()
            if not key or key in seen:
                continue
            seen.add(key)
            vendor = str(r["_vendor"]).strip()
            rows.append({
                "Vendor Order": order,
                "Vendor": vendor,
                "Vendor Code": key,
                "Vendor Key": key,
                "Source PSW File Order": file_order,
                "Rows": int((tmp["_vendor_key"] == key).sum()),
            })
            order += 1
    return pd.DataFrame(rows, columns=["Vendor Order", "Vendor", "Vendor Code", "Vendor Key", "Source PSW File Order", "Rows"])


def build_vendor_offset_maps(
    timeline_vendor_df: pd.DataFrame,
    due_date_calc_paths: List[str],
    ) -> Tuple[Dict[str, Dict[str, int]], pd.DataFrame]:
    """Build vendor -> warehouse offset map from DueDateCalc upload order.

    Rules:
    - If only one DueDateCalc file is uploaded, all vendors use that file.
    - If multiple files are uploaded, file order follows the detected PSW vendor order.
    - If fewer files than vendors are uploaded, the last uploaded file is reused as fallback.
    """
    due_paths = [p for p in (due_date_calc_paths or []) if p]
    vendor_maps: Dict[str, Dict[str, int]] = {}
    debug_rows = []

    if timeline_vendor_df is None or timeline_vendor_df.empty:
        return vendor_maps, pd.DataFrame([
            ["Vendor DueDate mapping", "No vendor detected from PSW; default DueDateCalc will be used"]
        ], columns=["Field", "Value"])

    if not due_paths:
        return vendor_maps, pd.DataFrame([
            ["Vendor DueDate mapping", "No DueDateCalc files provided for vendor mapping"]
        ], columns=["Field", "Value"])

    cache: Dict[str, Dict[str, int]] = {}
    for _, r in timeline_vendor_df.iterrows():
        vendor_key = str(r.get("Vendor Key") or r.get("Vendor Code") or "").strip()
        vendor = str(r.get("Vendor", vendor_key)).strip()
        order_val = int(r.get("Vendor Order", 1)) if str(r.get("Vendor Order", "")).strip() else 1
        due_idx = min(max(order_val - 1, 0), len(due_paths) - 1)
        due_path = due_paths[due_idx]

        if due_path not in cache:
            cache[due_path], _ = load_due_date_offsets(due_path)
        vendor_maps[vendor_key] = cache[due_path]
        debug_rows.append({
            "Vendor Order": order_val,
            "Vendor": vendor,
            "Vendor Code": vendor_key,
            "DueDateCalc Used": os.path.basename(due_path),
            "Mapping Rule": "same file for all vendors" if len(due_paths) == 1 else ("upload-order match" if order_val <= len(due_paths) else "last-file fallback"),
        })

    return vendor_maps, pd.DataFrame(debug_rows)


def find_transit_weeks_in_row(row: pd.Series, default_weeks: int) -> int:
    """
    Optional vendor-specific transit support.
    If a PSW/export row contains Transit Weeks, Transit Days, Delivery Days, or Lead Time columns,
    use that value. Otherwise fall back to the warehouse offset from DueDateCalc.
    """
    week_keywords = ["transit week", "delivery week", "lead week", "offset week"]
    day_keywords = ["transit day", "delivery day", "lead day"]
    for c in row.index:
        name = str(c).strip().lower()
        if any(k in name for k in week_keywords):
            val = pd.to_numeric(row.get(c), errors="coerce")
            if pd.notna(val):
                return max(0, int(math.ceil(float(val))))
    for c in row.index:
        name = str(c).strip().lower()
        if any(k in name for k in day_keywords):
            val = pd.to_numeric(row.get(c), errors="coerce")
            if pd.notna(val):
                return max(0, int(math.ceil(float(val) / 7.0)))
    return int(default_weeks)


def load_psw_vendor_supply(
    psw_csv_paths: List[str],
    target_week: date,
    current_week: date,
    offset_map: Dict[str, int],
    other_vendor_offset_map: Optional[Dict[str, int]] = None,
    vendor_offset_maps: Optional[Dict[str, Dict[str, int]]] = None,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Read one or more PSW / Production Schedule CSV files and return vendor-level F rows by week.

    PSW week is ETD. PlanDetailTimeline week is ETA. Adjusted Supply Week is kept for audit:
      Adjusted Supply Week = PSW ETD Week + Vendor Transit Weeks - Warehouse Offset Weeks

    If no vendor-specific transit column exists in PSW:
      - first PSW file uses the main/default DueDateCalc transit by warehouse;
      - second/subsequent PSW files use the optional sub-vendor DueDateCalc transit by warehouse;
      - if no sub-vendor DueDateCalc is provided, sub vendors fall back to the main/default transit.

    Warehouse Offset Weeks is always the main/default PlanDetailTimeline offset basis.
    Vendor Transit Weeks may differ by upload order, so sub-vendor supply can shift into an adjusted week.
    """
    if other_vendor_offset_map is None:
        other_vendor_offset_map = offset_map
    vendor_offset_maps = vendor_offset_maps or {}

    def _vendor_specific_transit(row, fallback_weeks):
        wh = normalize_whse(row.get("Whse", ""))
        vk = vendor_match_key(row.get("Vendor", ""))
        if vk in vendor_offset_maps and wh in vendor_offset_maps[vk]:
            return int(vendor_offset_maps[vk][wh])
        return int(fallback_weeks)

    detail_frames = []
    debug_rows = []
    for file_order, path in enumerate(psw_csv_paths or [], start=1):
        if not path:
            continue
        source_vendor_role = "MAIN_FILE" if file_order == 1 else "OTHER_FILE"
        prod = read_report_csv(path, dtype=str)
        prod.columns = [str(c).strip() for c in prod.columns]
        required = ["Item #", "Whse", "S/F/P"]
        missing = [c for c in required if c not in prod.columns]
        if missing:
            raise ValueError(f"PSW/Production Schedule is missing required columns {missing}: {path}")
        vendor_col = find_vendor_col(prod)
        if vendor_col is None:
            # Keep the flow working; vendor will be blank and all supply falls back to legacy item+whse logic.
            prod["Vendor"] = ""
            vendor_col = "Vendor"

        report_dt = extract_report_date_from_csv(path)
        date_map = build_production_date_map(prod.columns, report_dt, target_week)
        week_cols = sorted(date_map.items(), key=lambda x: x[0])

        prod = prod.copy()
        prod["Item"] = prod["Item #"].map(normalize_item)
        prod["Whse"] = prod["Whse"].map(normalize_whse)
        prod["Vendor"] = prod[vendor_col].map(normalize_vendor)
        prod["S/F/P"] = prod["S/F/P"].fillna("").astype(str).str.strip().str.upper()
        f = prod[prod["S/F/P"] == "F"].copy()

        rows = []
        for wk, col in week_cols:
            qty = pd.to_numeric(f[col], errors="coerce").fillna(0.0)
            nonzero = f.loc[qty != 0].copy()
            if nonzero.empty:
                continue
            nonzero["PSW Week"] = wk
            nonzero["PSW Week Text"] = fmt_date(wk)
            nonzero["PSW Quantity"] = qty.loc[qty != 0].values
            # Warehouse Offset Weeks is the PlanDetailTimeline/default ETD basis.
            nonzero["Warehouse Offset Weeks"] = nonzero["Whse"].map(offset_map).fillna(0).astype(int)

            # Keep both main/default and sub/other transit candidates.
            # The final vendor role is decided after matching to PlanDetailTimeline vendor.
            # This is important when one PSW file contains both main and sub vendors.
            nonzero["Main Default Vendor Transit Weeks"] = nonzero["Whse"].map(offset_map).fillna(nonzero["Warehouse Offset Weeks"]).astype(int)
            nonzero["Sub Default Vendor Transit Weeks"] = nonzero["Whse"].map(other_vendor_offset_map).fillna(nonzero["Main Default Vendor Transit Weeks"]).astype(int)

            # Temporary values for audit before role split; split_main_other_vendor_supply recomputes these
            # using final MAIN vs OTHER role, but vendor-specific DueDateCalc is already available here.
            initial_transit_map = offset_map if file_order == 1 else other_vendor_offset_map
            nonzero["Default Vendor Transit Weeks"] = nonzero["Whse"].map(initial_transit_map).fillna(nonzero["Warehouse Offset Weeks"]).astype(int)
            nonzero["Vendor Specific DueDateCalc Transit Weeks"] = [
                _vendor_specific_transit(r, int(r["Default Vendor Transit Weeks"])) for _, r in nonzero.iterrows()
            ]
            nonzero["Vendor Transit Source"] = "Vendor-specific DueDateCalc by PSW vendor order"
            nonzero["Vendor Transit Weeks"] = [
                find_transit_weeks_in_row(r, int(r["Vendor Specific DueDateCalc Transit Weeks"])) for _, r in nonzero.iterrows()
            ]
            nonzero["Adjusted Supply Week"] = [
                wk + timedelta(days=7 * (int(vt) - int(wo)))
                for vt, wo in zip(nonzero["Vendor Transit Weeks"], nonzero["Warehouse Offset Weeks"])
            ]
            nonzero["Adjusted Supply Week Text"] = nonzero["Adjusted Supply Week"].map(fmt_date)
            nonzero["Source File"] = os.path.basename(path)
            nonzero["Source File Order"] = file_order
            nonzero["Source Vendor Role"] = source_vendor_role
            rows.append(nonzero[[
                "Source File", "Source File Order", "Source Vendor Role", "Item", "Whse", "Vendor", "PSW Week", "PSW Week Text", "PSW Quantity",
                "Vendor Transit Source", "Default Vendor Transit Weeks", "Main Default Vendor Transit Weeks", "Sub Default Vendor Transit Weeks", "Vendor Specific DueDateCalc Transit Weeks", "Vendor Transit Weeks", "Warehouse Offset Weeks", "Adjusted Supply Week", "Adjusted Supply Week Text"
            ]])
        detail = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame(columns=[
            "Source File", "Source File Order", "Source Vendor Role", "Item", "Whse", "Vendor", "PSW Week", "PSW Week Text", "PSW Quantity",
            "Vendor Transit Source", "Default Vendor Transit Weeks", "Vendor Transit Weeks", "Warehouse Offset Weeks", "Adjusted Supply Week", "Adjusted Supply Week Text"
        ])
        detail_frames.append(detail)
        debug_rows.extend([
            [os.path.basename(path), "Source file order", file_order],
            [os.path.basename(path), "Source vendor role from upload order", source_vendor_role],
            [os.path.basename(path), "Rows", len(prod)],
            [os.path.basename(path), "F rows", len(f)],
            [os.path.basename(path), "Vendor column", vendor_col],
            [os.path.basename(path), "Week columns", len(week_cols)],
            [os.path.basename(path), "Nonzero F vendor-week rows", len(detail)],
            [os.path.basename(path), "Total nonzero F quantity", float(detail["PSW Quantity"].sum()) if not detail.empty else 0.0],
        ])
    all_detail = pd.concat(detail_frames, ignore_index=True) if detail_frames else pd.DataFrame(columns=[
        "Source File", "Source File Order", "Source Vendor Role", "Item", "Whse", "Vendor", "PSW Week", "PSW Week Text", "PSW Quantity",
        "Vendor Transit Source", "Default Vendor Transit Weeks", "Vendor Transit Weeks", "Warehouse Offset Weeks", "Adjusted Supply Week", "Adjusted Supply Week Text"
    ])
    debug = pd.DataFrame(debug_rows, columns=["Source File", "Field", "Value"])
    if debug.empty:
        debug = pd.DataFrame([["", "PSW files", 0]], columns=["Source File", "Field", "Value"])
    return all_detail, debug


def split_main_other_vendor_supply(
    base_rows: pd.DataFrame,
    psw_detail: pd.DataFrame,
    target_week: date,
    current_week: date,
    vendor_offset_maps: Optional[Dict[str, Dict[str, int]]] = None,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Split PSW supply into main-vendor and other-vendor buckets.
    Preferred rule: upload order decides role (1st PSW = main vendor, 2nd+ PSW = other vendor).
    Fallback rule: if upload-order role is unavailable, compare PSW vendor to PlanDetailTimeline vendor.
    - Main vendor: only Target Week quantity becomes F Wk3 for optimizer allocation.
    - Other vendor: uses the same Target Week PSW F logic as main vendor for subvendor suggestion; SI baseline is main-vendor SI After.
    """
    vendor_offset_maps = vendor_offset_maps or {}
    base = base_rows[["Item", "Whse", "Vendor"]].copy()
    base["Item"] = base["Item"].map(normalize_item)
    base["Whse"] = base["Whse"].map(normalize_whse)
    base["Main Vendor"] = base["Vendor"].map(normalize_vendor)
    main_map = base.groupby(["Item", "Whse"], dropna=False)["Main Vendor"].first().reset_index()

    if psw_detail is None or psw_detail.empty:
        empty_group = base[["Item", "Whse"]].drop_duplicates().copy()
        empty_group["Main Vendor F Wk3"] = 0.0
        empty_group["Other Vendor Supply"] = 0.0
        empty_group["Other Vendor List"] = ""
        empty_group["Total Supply Added to SI"] = 0.0
        return empty_group, pd.DataFrame(), pd.DataFrame()

    detail = psw_detail.copy()
    detail["Item"] = detail["Item"].map(normalize_item)
    detail["Whse"] = detail["Whse"].map(normalize_whse)
    detail["Vendor"] = detail["Vendor"].map(normalize_vendor)
    detail = detail.merge(main_map, on=["Item", "Whse"], how="left")

    # Vendor role rule:
    #   - If PSW vendor matches PlanDetailTimeline vendor, treat it as MAIN.
    #   - If PSW vendor does not match Timeline vendor, treat it as OTHER, even when it is in the first uploaded PSW file.
    #   - If the row comes from a second/subsequent PSW file, keep it as OTHER.
    # This handles the common case where one PSW file contains both the main vendor and sub vendor.
    has_source_role = "Source Vendor Role" in detail.columns and detail["Source Vendor Role"].fillna("").astype(str).str.strip().ne("").any()
    def _decide_vendor_role(r):
        source_role = str(r.get("Source Vendor Role", "")).upper().strip()
        psw_vendor_key = vendor_match_key(r.get("Vendor"))
        main_vendor_key = vendor_match_key(r.get("Main Vendor"))
        if has_source_role and source_role != "MAIN_FILE":
            return "OTHER"
        if main_vendor_key:
            return "MAIN" if psw_vendor_key == main_vendor_key else "OTHER"
        # Backward compatibility: if Timeline vendor is blank, first uploaded PSW file is main; later files are other.
        return "MAIN" if (not has_source_role or source_role == "MAIN_FILE") else "OTHER"

    detail["Vendor Role"] = detail.apply(_decide_vendor_role, axis=1)

    # Recompute transit/adjusted week after final role split.
    # MAIN rows use main/default DueDateCalc; OTHER rows use sub/other DueDateCalc if uploaded.
    if "Main Default Vendor Transit Weeks" not in detail.columns:
        detail["Main Default Vendor Transit Weeks"] = detail["Default Vendor Transit Weeks"]
    if "Sub Default Vendor Transit Weeks" not in detail.columns:
        detail["Sub Default Vendor Transit Weeks"] = detail["Default Vendor Transit Weeks"]
    detail["Default Vendor Transit Weeks"] = detail.apply(
        lambda r: r["Main Default Vendor Transit Weeks"] if r["Vendor Role"] == "MAIN" else r["Sub Default Vendor Transit Weeks"],
        axis=1,
    )

    def _role_vendor_transit(row):
        wh = normalize_whse(row.get("Whse", ""))
        vk = vendor_match_key(row.get("Vendor", ""))
        if vk in vendor_offset_maps and wh in vendor_offset_maps[vk]:
            return int(vendor_offset_maps[vk][wh]), "Vendor-specific DueDateCalc by PSW vendor order"
        return int(row["Default Vendor Transit Weeks"]), ("Main DueDateCalc fallback" if row["Vendor Role"] == "MAIN" else "Sub/default DueDateCalc fallback")

    _vt_pairs = [_role_vendor_transit(r) for _, r in detail.iterrows()]
    detail["Vendor Specific DueDateCalc Transit Weeks"] = [p[0] for p in _vt_pairs]
    detail["Vendor Transit Source"] = [p[1] for p in _vt_pairs]
    detail["Vendor Transit Weeks"] = [
        find_transit_weeks_in_row(r, int(r["Vendor Specific DueDateCalc Transit Weeks"])) for _, r in detail.iterrows()
    ]
    detail["Adjusted Supply Week"] = [
        pd.to_datetime(wk).date() + timedelta(days=7 * (int(vt) - int(wo)))
        for wk, vt, wo in zip(detail["PSW Week"], detail["Vendor Transit Weeks"], detail["Warehouse Offset Weeks"])
    ]
    detail["Adjusted Supply Week Text"] = detail["Adjusted Supply Week"].map(fmt_date)

    detail["Included as Main F Wk3"] = (detail["Vendor Role"] == "MAIN") & (detail["PSW Week"] == target_week)

    # Sub/other vendor must follow the same week logic as the main vendor.
    # The only difference is the SI baseline used later: subvendor DC starts from main-vendor SI After.
    # Therefore subvendor F Original / Other Vendor Supply is also PSW F at Target Week,
    # not Adjusted Supply Week and not Current Week -> Target Week window.
    detail["Included as Other F Wk3"] = (detail["Vendor Role"] == "OTHER") & (detail["PSW Week"] == target_week)
    detail["Included as Other Supply"] = detail["Included as Other F Wk3"]

    detail["Inclusion Reason"] = "Not included"
    detail.loc[detail["Included as Main F Wk3"], "Inclusion Reason"] = "Main vendor, PSW ETD week = Target Week; used as F Wk3 for optimizer"
    detail.loc[detail["Included as Other Supply"], "Inclusion Reason"] = "Other vendor, PSW ETD week = Target Week; used as subvendor F suggestion and SI/SS supply only"

    main = detail[detail["Included as Main F Wk3"]].groupby(["Item", "Whse"], dropna=False, as_index=False)["PSW Quantity"].sum().rename(columns={"PSW Quantity": "Main Vendor F Wk3"})
    other_qty = detail[detail["Included as Other Supply"]].groupby(["Item", "Whse"], dropna=False, as_index=False)["PSW Quantity"].sum().rename(columns={"PSW Quantity": "Other Vendor Supply"})

    # Show the other/sub vendor code on every warehouse row of the same item for easier filtering/debugging.
    # Quantity is still counted only on rows where the other vendor has PSW F at Target Week.
    item_other_list = (
        detail[detail["Vendor Role"] == "OTHER"]
        .groupby(["Item"], dropna=False)["Vendor"]
        .agg(lambda x: ", ".join(sorted({str(v).strip() for v in x if str(v).strip() and str(v).strip().lower() != "nan"})))
        .reset_index()
        .rename(columns={"Vendor": "Other Vendor List"})
    )

    grouped = (
        base[["Item", "Whse", "Main Vendor"]].drop_duplicates()
        .merge(main, on=["Item", "Whse"], how="left")
        .merge(other_qty, on=["Item", "Whse"], how="left")
        .merge(item_other_list, on=["Item"], how="left")
    )
    grouped["Main Vendor F Wk3"] = pd.to_numeric(grouped["Main Vendor F Wk3"], errors="coerce").fillna(0.0)
    grouped["Other Vendor Supply"] = pd.to_numeric(grouped["Other Vendor Supply"], errors="coerce").fillna(0.0)
    grouped["Other Vendor List"] = grouped["Other Vendor List"].fillna("")
    # Initial PSW supply only. Final Total Supply Added to SI is recomputed after Timeline Firm PO is available
    # as Main Vendor F Wk3 + Other Vendor Supply + Firm PO Reconciliation Gap.
    grouped["PSW F Used for Reconciliation"] = grouped["Main Vendor F Wk3"] + grouped["Other Vendor Supply"]
    grouped["Total Supply Added to SI"] = grouped["PSW F Used for Reconciliation"]

    supply_debug = pd.DataFrame([
        ["PSW vendor rows", len(detail)],
        ["Main vendor rows included as F Wk3", int(detail["Included as Main F Wk3"].sum())],
        ["Other vendor rows included as SI/SS supply", int(detail["Included as Other Supply"].sum())],
        ["Main Vendor F Wk3 total", float(grouped["Main Vendor F Wk3"].sum())],
        ["Other Vendor Supply total", float(grouped["Other Vendor Supply"].sum())],
        ["Total Supply Added to SI", float(grouped["Total Supply Added to SI"].sum())],
        ["PSW role rule", "Vendor match to PlanDetailTimeline decides MAIN vs OTHER inside each PSW file; second/subsequent PSW files are forced OTHER"],
        ["Other vendor inclusion rule", "Same as main vendor: PSW ETD week = Target Week. Adjusted Supply Week is audit only."],
    ], columns=["Field", "Value"])
    return grouped, detail, supply_debug

def load_fwk3_from_production(production_csv_path: str, target_week: date) -> Tuple[pd.DataFrame, pd.DataFrame]:
    prod = read_report_csv(production_csv_path, dtype=str)
    prod.columns = [str(c).strip() for c in prod.columns]
    required = ["Item #", "Whse", "S/F/P"]
    missing = [c for c in required if c not in prod.columns]
    if missing:
        raise ValueError(f"Production Schedule is missing required columns: {missing}")

    report_dt = extract_report_date_from_csv(production_csv_path)
    date_map = build_production_date_map(prod.columns, report_dt, target_week)
    week_col = date_map.get(target_week)
    if week_col is None:
        available = ", ".join(fmt_date(d) for d in sorted(date_map.keys())[:5]) + " ... " + ", ".join(fmt_date(d) for d in sorted(date_map.keys())[-5:])
        raise ValueError(f"Production Schedule does not contain the target week column {fmt_date(target_week)}. Available: {available}")

    prod = prod.copy()
    prod["Item"] = prod["Item #"].map(normalize_item)
    prod["Whse"] = prod["Whse"].map(normalize_whse)
    prod["S/F/P"] = prod["S/F/P"].fillna("").astype(str).str.strip().str.upper()
    prod["F Wk3"] = pd.to_numeric(prod[week_col], errors="coerce").fillna(0.0)

    f = prod[prod["S/F/P"] == "F"].copy()
    grouped = f.groupby(["Item", "Whse"], dropna=False, as_index=False)["F Wk3"].sum()

    debug = pd.DataFrame([
        ["Production rows", len(prod)],
        ["Production F rows", len(f)],
        ["Production report date", str(report_dt) if report_dt else ""],
        ["TargetWeek", fmt_date(target_week)],
        ["Production target column", week_col],
        ["F Wk3 total", float(grouped["F Wk3"].sum())],
        ["F Wk3 nonzero item-whse", int((grouped["F Wk3"] != 0).sum())],
    ], columns=["Field", "Value"])
    return grouped, debug




def build_optimizer_input_direct_from_plan(
    plan_csv_path: str,
    offset_map: Dict[str, int],
    f_wk3: pd.DataFrame,
    target_week: date,
    current_week: date,
    psw_supply_detail: Optional[pd.DataFrame] = None,
    vendor_offset_maps: Optional[Dict[str, Dict[str, int]]] = None,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Fast path: compute optimizer input directly from raw PlanDetailTimeline without materializing 44-week converted file."""
    raw = read_report_csv(plan_csv_path, dtype=str)
    raw.columns = [str(c).strip() for c in raw.columns]
    required_base = ["Item #", "Whse", "Data Type", "Coll. Class", "MakeBuy Code"]
    missing = [c for c in required_base if c not in raw.columns]
    if missing:
        raise ValueError(f"PlanDetailTimeline is missing required columns: {missing}")

    raw = raw.copy()
    raw["Item #"] = raw["Item #"].map(normalize_item)
    raw["Whse"] = raw["Whse"].map(normalize_whse)
    raw["Data Type"] = clean_dtype(raw["Data Type"])
    raw["MakeBuy Code"] = raw["MakeBuy Code"].fillna("").astype(str).str.strip().str.upper()
    raw["Coll. Class"] = raw["Coll. Class"].fillna("").astype(str).str.strip()

    timeline_vendor_col = find_vendor_col(raw)
    if timeline_vendor_col:
        raw["_timeline_vendor"] = raw[timeline_vendor_col].map(normalize_vendor)
        raw["_timeline_vendor_key"] = raw["_timeline_vendor"].map(vendor_match_key)
    else:
        raw["_timeline_vendor"] = ""
        raw["_timeline_vendor_key"] = ""

    # Robustly detect weekly date columns by header value.
    # Do not rely on fixed column positions because PlanDetailTimeline exports may
    # include a different number of trailing attributes.
    date_cols = []
    date_map = {}
    for c in raw.columns:
        d = parse_header_to_date(c)
        if d is not None:
            date_cols.append(c)
            date_map[d] = c
    if not date_cols:
        raise ValueError("PlanDetailTimeline does not contain recognizable weekly date columns.")
    original_dates = sorted(date_map.keys())
    first_original_week = min(original_dates)
    last_original_week = max(original_dates)
    first_etd_week = first_original_week - timedelta(days=154)

    raw = raw[raw["MakeBuy Code"] == "B"].copy()
    if raw.empty:
        raise ValueError("No rows remain after filtering MakeBuy Code = B.")

    raw["_target_value"] = 0.0
    raw["_planned_sum"] = 0.0
    raw["_net_sum"] = 0.0

    vendor_offset_maps = vendor_offset_maps or {}
    def _offset_for_row(row):
        wh = normalize_whse(row.get("Whse", ""))
        vk = str(row.get("_timeline_vendor_key", "")).strip()
        if vk in vendor_offset_maps and wh in vendor_offset_maps[vk]:
            return int(vendor_offset_maps[vk][wh])
        return int(offset_map.get(wh, 0))

    raw["_offset_weeks"] = raw.apply(_offset_for_row, axis=1).astype(int)
    raw["_offset_source"] = raw.apply(
        lambda r: f"Vendor DueDateCalc {r.get('_timeline_vendor_key','')}"
        if str(r.get("_timeline_vendor_key", "")).strip() in vendor_offset_maps
        else "Default DueDateCalc",
        axis=1,
    )
    unknown_whse = sorted(set(raw.loc[~raw["Whse"].isin(offset_map.keys()), "Whse"].dropna().astype(str)))

    planned_etd_weeks = date_range_saturdays(first_etd_week, target_week)
    net_etd_weeks = date_range_saturdays(current_week, target_week)
    if current_week > target_week:
        raise ValueError("Target Week phai lon hon hoac bang Current Week.")

    # Process by offset instead of by row, much faster.
    for offset, idx in raw.groupby("_offset_weeks", sort=False).groups.items():
        target_src = target_week + timedelta(days=7 * int(offset))
        target_col = date_map.get(target_src)
        if target_col:
            raw.loc[idx, "_target_value"] = pd.to_numeric(raw.loc[idx, target_col], errors="coerce").fillna(0.0).values

        planned_cols = [date_map[d + timedelta(days=7 * int(offset))] for d in planned_etd_weeks if (d + timedelta(days=7 * int(offset))) in date_map]
        if planned_cols:
            raw.loc[idx, "_planned_sum"] = raw.loc[idx, planned_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0).sum(axis=1).values

        direct_net_cols = [date_map[d] for d in net_etd_weeks if d in date_map]
        if direct_net_cols:
            rows_idx = raw.loc[idx]
            is_335 = rows_idx["Whse"].astype(str).eq("335")
            if is_335.any():
                rows_335 = rows_idx.loc[is_335]
                raw.loc[rows_335.index, "_net_sum"] = rows_335[direct_net_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0).sum(axis=1).values
            if (~is_335).any():
                rows_other = rows_idx.loc[~is_335]
                shifted_net_cols = [date_map[d + timedelta(days=7 * int(offset))] for d in net_etd_weeks if d + timedelta(days=7 * int(offset)) in date_map]
                if shifted_net_cols:
                    raw.loc[rows_other.index, "_net_sum"] = rows_other[shifted_net_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0).sum(axis=1).values

    key_cols = ["Item #", "Whse", "Coll. Class"]
    si_g = group_value(raw[raw["Data Type"] == "SHIPPABLE INV"], key_cols, "_target_value", "Base_SI")
    planned_g = group_value(raw[raw["Data Type"] == "PLANNED POS"], key_cols, "_planned_sum", "PlannedPO_Sum")
    firm_g = group_value(raw[raw["Data Type"] == "FIRM POS"], key_cols, "_target_value", "FirmPO_Target")
    net_g = group_value(raw[raw["Data Type"] == "NET FCST"], key_cols, "_net_sum", "NetFcst_Sum")
    ss_g = group_value(raw[raw["Data Type"] == "SAFETY STK"], key_cols, "_target_value", "SS_Wk3")

    base = raw[raw["Data Type"].isin(["SHIPPABLE INV", "PLANNED POS", "FIRM POS", "NET FCST", "SAFETY STK"])][key_cols].drop_duplicates()
    out = base.merge(si_g, on=key_cols, how="left").merge(planned_g, on=key_cols, how="left").merge(firm_g, on=key_cols, how="left").merge(net_g, on=key_cols, how="left").merge(ss_g, on=key_cols, how="left")
    for col in ["Base_SI", "PlannedPO_Sum", "FirmPO_Target", "NetFcst_Sum", "SS_Wk3"]:
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0.0)

    # Timeline Firm PO at the mapped ETA week. It is kept for reconciliation because
    # Current SI has already subtracted this Firm PO amount from PlanDetailTimeline.
    out["Timeline Firm PO"] = out["FirmPO_Target"]

    out["F Wk3"] = 0.0
    # Standard SI basis. Warehouse 335 has a special rule: add accumulated NET FCST
    # from Current Week through the selected balancing/target week.
    out["Sum of SI Wk3"] = out["Base_SI"] - out["PlannedPO_Sum"] - out["FirmPO_Target"]
    mask_335 = out["Whse"].astype(str) == "335"
    out.loc[mask_335, "Sum of SI Wk3"] = (
        out.loc[mask_335, "Base_SI"]
        - out.loc[mask_335, "PlannedPO_Sum"]
        - out.loc[mask_335, "FirmPO_Target"]
        + out.loc[mask_335, "NetFcst_Sum"]
    )
    out["Sum of SI-SS Wk3"] = out["Sum of SI Wk3"] - out["SS_Wk3"]
    out["Average of SS Wk3"] = out["SS_Wk3"]

    vendor_col = next((c for c in raw.columns if c.lower() == "vendor"), None)
    if vendor_col:
        vendor_map = raw.groupby(key_cols, dropna=False)[vendor_col].first().reset_index().rename(columns={vendor_col: "Vendor"})
        out = out.merge(vendor_map, on=key_cols, how="left")
    else:
        out["Vendor"] = ""

    out = out.rename(columns={"Item #": "Item", "Coll. Class": "ProdResourceID"})
    out["Item"] = out["Item"].map(normalize_item)
    out["Whse"] = out["Whse"].map(normalize_whse)

    # Multi-vendor PSW logic.
    # Vendor matching rule:
    #   PSW Vendor == Timeline Vendor -> main vendor. Main vendor Target Week quantity becomes F Wk3.
    #   PSW Vendor != Timeline Vendor -> other vendor. Other vendor quantity updates SI/SS only, not F Wk3 allocation.
    if psw_supply_detail is not None and not psw_supply_detail.empty:
        supply_grouped, psw_vendor_detail, psw_supply_debug = split_main_other_vendor_supply(
            out[["Item", "Whse", "Vendor"]].drop_duplicates(), psw_supply_detail, target_week, current_week,
            vendor_offset_maps=vendor_offset_maps,
        )
        out = out.merge(supply_grouped, on=["Item", "Whse"], how="left")
        out["Main Vendor"] = out["Main Vendor"].fillna(out["Vendor"].map(normalize_vendor))
        out["Main Vendor F Wk3"] = pd.to_numeric(out["Main Vendor F Wk3"], errors="coerce").fillna(0.0)
        out["Other Vendor Supply"] = pd.to_numeric(out["Other Vendor Supply"], errors="coerce").fillna(0.0)
        out["Other Vendor List"] = out["Other Vendor List"].fillna("")
        out["PSW F Used for Reconciliation"] = out["Main Vendor F Wk3"] + out["Other Vendor Supply"]
        # Firm PO Reconciliation Gap = Timeline Firm PO at mapped ETA week - PSW F around mapped ETD bucket.
        # This gap is added back to New SI / New SI-SS because Timeline Current SI already used Timeline Firm PO.
        out["Firm PO Reconciliation Gap"] = out["Timeline Firm PO"] - out["PSW F Used for Reconciliation"]
        out["Total Supply Added to SI"] = (
            out["Main Vendor F Wk3"] + out["Other Vendor Supply"] + out["Firm PO Reconciliation Gap"]
        )
        out["F Wk3"] = out["Main Vendor F Wk3"]
        f_for_missing = supply_grouped[["Item", "Whse", "Main Vendor F Wk3"]].rename(columns={"Main Vendor F Wk3": "F Wk3"})
    else:
        psw_vendor_detail = pd.DataFrame()
        psw_supply_debug = pd.DataFrame([["PSW vendor detail", "Not provided; using default item+warehouse F Wk3"]], columns=["Field", "Value"])
        f_wk3 = f_wk3.copy()
        f_wk3["Item"] = f_wk3["Item"].map(normalize_item)
        f_wk3["Whse"] = f_wk3["Whse"].map(normalize_whse)
        out = out.merge(f_wk3, on=["Item", "Whse"], how="left", suffixes=("", "_from_prod"))
        firm_col = next((c for c in ["F Wk3_from_prod", "F Wk3_y", "F Wk3"] if c in out.columns), None)
        if firm_col is None:
            out["F Wk3"] = 0.0
        else:
            if firm_col != "F Wk3":
                out["F Wk3"] = pd.to_numeric(out[firm_col], errors="coerce").fillna(0.0)
                out = out.drop(columns=[firm_col])
            else:
                out["F Wk3"] = pd.to_numeric(out["F Wk3"], errors="coerce").fillna(0.0)
        out["Main Vendor"] = out["Vendor"].map(normalize_vendor)
        out["Main Vendor F Wk3"] = out["F Wk3"]
        out["Other Vendor Supply"] = 0.0
        out["Other Vendor List"] = ""
        out["PSW F Used for Reconciliation"] = out["Main Vendor F Wk3"] + out["Other Vendor Supply"]
        out["Firm PO Reconciliation Gap"] = out["Timeline Firm PO"] - out["PSW F Used for Reconciliation"]
        out["Total Supply Added to SI"] = (
            out["Main Vendor F Wk3"] + out["Other Vendor Supply"] + out["Firm PO Reconciliation Gap"]
        )
        f_for_missing = f_wk3

    # Preserve exactly the approved 9 PlanDetailTimeline metadata columns.
    # Resolve them at Item + Whse level using the first nonblank value in the B-scope rows.
    meta_cols = [c for c in PLAN_METADATA_COLUMNS if c in raw.columns]
    if meta_cols:
        meta_src = raw[["Item #", "Whse"] + meta_cols].copy()
        for c in meta_cols:
            meta_src[c] = meta_src[c].fillna("").astype(str).str.strip()
        def _first_nonblank(series):
            vals = series[series.astype(str).str.strip().ne("") & series.astype(str).str.lower().ne("nan")]
            return vals.iloc[0] if not vals.empty else ""
        meta_map = meta_src.groupby(["Item #", "Whse"], dropna=False)[meta_cols].agg(_first_nonblank).reset_index()
        meta_map = meta_map.rename(columns={"Item #": "Item"})
        meta_map["Item"] = meta_map["Item"].map(normalize_item)
        meta_map["Whse"] = meta_map["Whse"].map(normalize_whse)
        out = out.merge(meta_map, on=["Item", "Whse"], how="left")

    output_cols = [c for c in OUTPUT_COLUMNS if c in out.columns]
    output = out[output_cols + [c for c in PLAN_METADATA_COLUMNS if c in out.columns]].drop_duplicates().copy()

    merge_debug = output.merge(f_for_missing, on=["Item", "Whse"], how="left", indicator=True, suffixes=("", "_prod"))
    missing_f = merge_debug[merge_debug["_merge"] == "left_only"][["Item", "Whse", "ProdResourceID"]].drop_duplicates()

    build_debug = pd.DataFrame([
        ["Plan rows after MakeBuy B", len(raw)],
        ["Original first ETA week", fmt_date(first_original_week)],
        ["Original last ETA week", fmt_date(last_original_week)],
        ["First converted ETD week", fmt_date(first_etd_week)],
        ["TargetWeek", fmt_date(target_week)],
        ["CurrentWeek", fmt_date(current_week)],
        ["Planned POS ETD range", ", ".join(fmt_date(d) for d in planned_etd_weeks)],
        ["NET FCST ETD range", ", ".join(fmt_date(d) for d in net_etd_weeks)],
        ["F Wk3 source", "PSW/Production Schedule: S/F/P = F, main vendor only, Target Week only"],
        ["Other Vendor Supply source", "PSW vendor different from Timeline vendor; adjusted supply week between Current Week and Target Week"],
        ["Rows output", str(len(output))],
        ["Rows without Production F match", str(len(missing_f))],
        ["F Wk3 total in optimizer input", str(float(output["F Wk3"].sum()))],
        ["Other Vendor Supply total", str(float(output["Other Vendor Supply"].sum())) if "Other Vendor Supply" in output.columns else "0"],
        ["Firm PO Reconciliation Gap total", str(float(output["Firm PO Reconciliation Gap"].sum())) if "Firm PO Reconciliation Gap" in output.columns else "0"],
        ["Total Supply Added to SI", str(float(output["Total Supply Added to SI"].sum())) if "Total Supply Added to SI" in output.columns else str(float(output["F Wk3"].sum()))],
        ["Unknown Whse count", len(unknown_whse)],
        ["Unknown Whse list", ", ".join(unknown_whse)],
        ["Whse 335 SI logic", "Same as other warehouses: SI(Target Week) - Planned POS(First ETD week -> Target Week) - Firm POS(Target Week). Net Forecast is not added."],
        ["Other Whse SI logic", "SI(Target Week) - Planned POS(First ETD week -> Target Week) - Firm POS(Target Week)"],
    ], columns=["Field", "Value"])

    offset_by_whse = raw[["Whse", "_timeline_vendor", "_timeline_vendor_key", "_offset_weeks", "_offset_source"]].drop_duplicates().rename(columns={"_timeline_vendor": "Timeline Vendor", "_timeline_vendor_key": "Vendor Code", "_offset_weeks": "Used Offset Weeks", "_offset_source": "Offset Source"}).sort_values(["Whse", "Vendor Code"])
    return output, build_debug, missing_f, offset_by_whse, psw_vendor_detail, psw_supply_debug


# ============================================================
# Step 4: Build optimizer input from converted PlanDetailTimeline
# ============================================================

def transform_converted_plan_to_optimizer_input(
    converted: pd.DataFrame,
    f_wk3: pd.DataFrame,
    target_week: date,
    current_week: date,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    raw = converted.copy()
    raw.columns = [str(c).strip() for c in raw.columns]

    required_base = ["Item #", "Whse", "Data Type", "Coll. Class", "MakeBuy Code"]
    missing_base = [c for c in required_base if c not in raw.columns]
    if missing_base:
        raise ValueError(f"Converted plan is missing required columns: {', '.join(missing_base)}")

    raw["Data Type"] = clean_dtype(raw["Data Type"])
    raw["MakeBuy Code"] = raw["MakeBuy Code"].fillna("").astype(str).str.strip().str.upper()
    raw["Item #"] = raw["Item #"].map(normalize_item)
    raw["Whse"] = raw["Whse"].map(normalize_whse)
    raw["Coll. Class"] = raw["Coll. Class"].fillna("").astype(str).str.strip()

    vendor_col = next((c for c in raw.columns if c.lower() == "vendor"), None)

    raw = raw[raw["MakeBuy Code"] == "B"].copy()
    if raw.empty:
        raise ValueError("No rows remain after filtering MakeBuy Code = B.")

    date_col_map = build_date_column_map(raw)
    target_col = date_col_map.get(target_week)
    if target_col is None:
        raise ValueError(f"Could not find the Target Week column in the converted plan: {fmt_date(target_week)}")

    all_week_dates = sorted(date_col_map.keys())
    first_week_date = min(all_week_dates)

    planned_weeks = date_range_saturdays(first_week_date, target_week)
    planned_missing = [fmt_date(d) for d in planned_weeks if d not in date_col_map]
    if planned_missing:
        raise ValueError("Missing weekly columns for Planned POS: " + ", ".join(planned_missing))

    if current_week > target_week:
        raise ValueError("Target Week phai lon hon hoac bang Current Week.")

    net_weeks = date_range_saturdays(current_week, target_week)
    net_missing = [fmt_date(d) for d in net_weeks if d not in date_col_map]
    if net_missing:
        raise ValueError("Missing weekly columns for NET FCST: " + ", ".join(net_missing))

    planned_cols = [date_col_map[d] for d in planned_weeks]
    net_cols = [date_col_map[d] for d in net_weeks]
    key_cols = ["Item #", "Whse", "Coll. Class"]

    si = raw[raw["Data Type"] == "SHIPPABLE INV"].copy()
    si["Base_SI"] = get_numeric(si, target_col)
    si_g = group_value(si, key_cols, "Base_SI", "Base_SI")

    planned = raw[raw["Data Type"] == "PLANNED POS"].copy()
    planned["PlannedPO_Sum"] = sum((get_numeric(planned, c) for c in planned_cols), start=pd.Series(0.0, index=planned.index))
    planned_g = group_value(planned, key_cols, "PlannedPO_Sum", "PlannedPO_Sum")

    firm = raw[raw["Data Type"] == "FIRM POS"].copy()
    firm["FirmPO_Target"] = get_numeric(firm, target_col)
    firm_g = group_value(firm, key_cols, "FirmPO_Target", "FirmPO_Target")

    net_fcst = raw[raw["Data Type"] == "NET FCST"].copy()
    net_fcst["NetFcst_Sum"] = sum((get_numeric(net_fcst, c) for c in net_cols), start=pd.Series(0.0, index=net_fcst.index))
    net_fcst_g = group_value(net_fcst, key_cols, "NetFcst_Sum", "NetFcst_Sum")

    ss = raw[raw["Data Type"] == "SAFETY STK"].copy()
    ss["SS_Wk3"] = get_numeric(ss, target_col)
    ss_g = group_value(ss, key_cols, "SS_Wk3", "SS_Wk3")

    base = raw[raw["Data Type"].isin(["SHIPPABLE INV", "PLANNED POS", "FIRM POS", "NET FCST", "SAFETY STK"])][key_cols].drop_duplicates()
    out = (
        base.merge(si_g, on=key_cols, how="left")
            .merge(planned_g, on=key_cols, how="left")
            .merge(firm_g, on=key_cols, how="left")
            .merge(net_fcst_g, on=key_cols, how="left")
            .merge(ss_g, on=key_cols, how="left")
    )

    for col in ["Base_SI", "PlannedPO_Sum", "FirmPO_Target", "NetFcst_Sum", "SS_Wk3"]:
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0.0)

    # Timeline Firm PO at the mapped ETA week. It is kept for reconciliation because
    # Current SI has already subtracted this Firm PO amount from PlanDetailTimeline.
    out["Timeline Firm PO"] = out["FirmPO_Target"]

    out["F Wk3"] = 0.0
    # SI logic applies the same formula for all warehouses, including Whse 335.
    # Net Forecast is no longer added for Whse 335.
    out["Sum of SI Wk3"] = out["Base_SI"] - out["PlannedPO_Sum"] - out["FirmPO_Target"]
    out["Sum of SI-SS Wk3"] = out["Sum of SI Wk3"] - out["SS_Wk3"]
    out["Average of SS Wk3"] = out["SS_Wk3"]

    if vendor_col:
        vendor_map = raw.groupby(["Item #", "Whse", "Coll. Class"], dropna=False)[vendor_col].first().reset_index().rename(columns={vendor_col: "Vendor"})
        out = out.merge(vendor_map, on=key_cols, how="left")
    else:
        out["Vendor"] = ""

    out = out.rename(columns={"Item #": "Item", "Coll. Class": "ProdResourceID"})
    out["Item"] = out["Item"].map(normalize_item)
    out["Whse"] = out["Whse"].map(normalize_whse)

    f_wk3 = f_wk3.copy()
    f_wk3["Item"] = f_wk3["Item"].map(normalize_item)
    f_wk3["Whse"] = f_wk3["Whse"].map(normalize_whse)
    out = out.merge(f_wk3, on=["Item", "Whse"], how="left", suffixes=("", "_from_prod"))
    firm_col = next((c for c in ["F Wk3_from_prod", "F Wk3_y", "F Wk3"] if c in out.columns), None)
    if firm_col is None:
        out["F Wk3"] = 0.0
    else:
        if firm_col != "F Wk3":
            out["F Wk3"] = pd.to_numeric(out[firm_col], errors="coerce").fillna(0.0)
            out = out.drop(columns=[firm_col])
        else:
            out["F Wk3"] = pd.to_numeric(out["F Wk3"], errors="coerce").fillna(0.0)
    out["Main Vendor"] = out["Vendor"].map(normalize_vendor)
    out["Main Vendor F Wk3"] = out["F Wk3"]
    out["Other Vendor Supply"] = 0.0
    out["Other Vendor List"] = ""
    out["PSW F Used for Reconciliation"] = out["Main Vendor F Wk3"] + out["Other Vendor Supply"]
    out["Firm PO Reconciliation Gap"] = out["Timeline Firm PO"] - out["PSW F Used for Reconciliation"]
    out["Total Supply Added to SI"] = (
        out["Main Vendor F Wk3"] + out["Other Vendor Supply"] + out["Firm PO Reconciliation Gap"]
    )

    output = out[[c for c in OUTPUT_COLUMNS if c in out.columns]].drop_duplicates().copy()

    merge_debug = output.merge(f_wk3, on=["Item", "Whse"], how="left", indicator=True, suffixes=("", "_prod"))
    missing_f = merge_debug[merge_debug["_merge"] == "left_only"][["Item", "Whse", "ProdResourceID"]].drop_duplicates()

    debug_rows = [
        ["First week in converted file", fmt_date(first_week_date)],
        ["TargetWeek", fmt_date(target_week)],
        ["CurrentWeek", fmt_date(current_week)],
        ["Planned POS range", ", ".join(fmt_date(d) for d in planned_weeks)],
        ["NET FCST range", ", ".join(fmt_date(d) for d in net_weeks)],
        ["TargetWeek column found", str(target_col)],
        ["F Wk3 source", "PSW/Production Schedule: S/F/P = F, main vendor only, Target Week only"],
        ["Other Vendor Supply source", "PSW vendor different from Timeline vendor; adjusted supply week between Current Week and Target Week"],
        ["Rows output", str(len(output))],
        ["Rows without Production F match", str(len(missing_f))],
        ["F Wk3 total in optimizer input", str(float(output["F Wk3"].sum()))],
        ["Other Vendor Supply total", str(float(output["Other Vendor Supply"].sum())) if "Other Vendor Supply" in output.columns else "0"],
        ["Firm PO Reconciliation Gap total", str(float(output["Firm PO Reconciliation Gap"].sum())) if "Firm PO Reconciliation Gap" in output.columns else "0"],
        ["Total Supply Added to SI", str(float(output["Total Supply Added to SI"].sum())) if "Total Supply Added to SI" in output.columns else str(float(output["F Wk3"].sum()))],
        ["Whse 335 SI logic", "SI(eval/balancing week) - Planned POS using ETA->ETD mapping - Firm POS at eval week + accumulated Net Forecast from Current Week to eval/balancing week without offset shift."],
        ["Other Whse SI logic", "SI(Target Week) - Planned POS(First week -> Target Week) - Firm POS(Target Week)"],
    ]
    debug_df = pd.DataFrame(debug_rows, columns=["Field", "Value"])
    return output, debug_df, missing_f


# ============================================================
# Auto balancing-week selection
# ============================================================

def load_plan_source(plan_csv_path: str) -> Tuple[pd.DataFrame, Dict[date, str], date, date, date]:
    raw = read_report_csv(plan_csv_path, dtype=str)
    raw.columns = [str(c).strip() for c in raw.columns]
    required = ["Item #", "Whse", "Data Type", "Coll. Class", "MakeBuy Code"]
    missing = [c for c in required if c not in raw.columns]
    if missing:
        raise ValueError(f"PlanDetailTimeline is missing required columns: {missing}")
    date_map = build_date_column_map(raw)
    if not date_map:
        raise ValueError("PlanDetailTimeline does not contain recognizable weekly date columns.")
    original_dates = sorted(date_map.keys())
    return raw, date_map, min(original_dates), max(original_dates), min(original_dates) - timedelta(days=154)


def compute_plan_metrics_for_week_from_raw(
    raw: pd.DataFrame,
    date_map: Dict[date, str],
    first_etd_week: date,
    eval_week: date,
    current_week: date,
    offset_map: Dict[str, int],
    firm_week: Optional[date] = None,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if current_week > eval_week:
        raise ValueError("Current Week cannot be later than the balancing week.")
    if firm_week is None:
        firm_week = eval_week
    if current_week > firm_week:
        raise ValueError("Current Week cannot be later than the Firm PO week.")
    df = raw.copy()
    df["Item #"] = df["Item #"].map(normalize_item)
    df["Whse"] = df["Whse"].map(normalize_whse)
    df["Data Type"] = clean_dtype(df["Data Type"])
    df["MakeBuy Code"] = df["MakeBuy Code"].fillna("").astype(str).str.strip().str.upper()
    df["Coll. Class"] = df["Coll. Class"].fillna("").astype(str).str.strip()
    df = df[df["MakeBuy Code"] == "B"].copy()
    if df.empty:
        raise ValueError("No rows remain after filtering MakeBuy Code = B.")

    df["_target_value"] = 0.0
    df["_planned_sum"] = 0.0
    df["_net_sum"] = 0.0
    df["_firm_week_value"] = 0.0
    df["_firm_through_balance_value"] = 0.0
    df["_offset_weeks"] = df["Whse"].map(offset_map).fillna(0).astype(int)
    # Safety Stock is special for Auto Separate / phase-out items.
    # Keep the raw value at the candidate week, but when it is <= 0 use the
    # most recent PRIOR positive Safety Stock for the same row. This is the
    # approved Last Positive SS rule and prevents phase-out rows from creating
    # INF / -INF merely because the current candidate SS dropped to zero.
    df["_ss_raw_value"] = 0.0
    df["_ss_effective_value"] = 0.0
    df["_ss_fallback_used"] = False
    # Number of consecutive candidate-side weekly buckets with SS <= 0, ending at eval_week.
    # 0 means current SS is positive; 1 means first zero-SS week; 2+ enters Runout Balance Mode.
    df["_ss_zero_streak"] = 0

    planned_weeks = date_range_saturdays(first_etd_week, eval_week)
    net_weeks = date_range_saturdays(current_week, eval_week)
    firm_eval_weeks = date_range_saturdays(firm_week, eval_week)
    for offset, idx in df.groupby("_offset_weeks", sort=False).groups.items():
        off = int(offset)
        target_src = eval_week + timedelta(days=7 * off)
        target_col = date_map.get(target_src)
        if target_col:
            df.loc[idx, "_target_value"] = pd.to_numeric(df.loc[idx, target_col], errors="coerce").fillna(0.0).values

        # Last Positive SS. The candidate's own mapped SS is used when > 0.
        # If it is 0 (or negative), walk backward week-by-week in the raw
        # PlanDetail timeline and take the nearest positive SS. Repeated zero
        # weeks therefore continue to use the same last positive value.
        ss_idx = [i for i in idx if str(df.at[i, "Data Type"]) == "SAFETY STK"]
        if ss_idx:
            if target_col:
                raw_ss = pd.to_numeric(df.loc[ss_idx, target_col], errors="coerce").fillna(0.0)
            else:
                raw_ss = pd.Series(0.0, index=ss_idx, dtype="float64")
            effective_ss = raw_ss.astype(float).copy()
            unresolved = effective_ss.le(0.0)
            if unresolved.any():
                prior_dates = sorted((d for d in date_map.keys() if d < target_src), reverse=True)
                for prior_date in prior_dates:
                    if not unresolved.any():
                        break
                    prior_col = date_map[prior_date]
                    prior_vals = pd.to_numeric(df.loc[ss_idx, prior_col], errors="coerce").fillna(0.0)
                    take = unresolved & prior_vals.gt(0.0)
                    if take.any():
                        effective_ss.loc[take] = prior_vals.loc[take]
                        unresolved = effective_ss.le(0.0)
            # Consecutive zero-SS streak ending at the mapped candidate week.
            zero_streak = pd.Series(0, index=ss_idx, dtype="int64")
            unresolved_zero = raw_ss.le(0.0)
            if unresolved_zero.any():
                zero_streak.loc[unresolved_zero] = 1
                prior_dates_asc = sorted((d for d in date_map.keys() if d < target_src), reverse=True)
                still_zero = unresolved_zero.copy()
                for prior_date in prior_dates_asc:
                    if not still_zero.any():
                        break
                    prior_col = date_map[prior_date]
                    prior_vals = pd.to_numeric(df.loc[ss_idx, prior_col], errors="coerce").fillna(0.0)
                    continued = still_zero & prior_vals.le(0.0)
                    if continued.any():
                        zero_streak.loc[continued] = zero_streak.loc[continued] + 1
                    still_zero = continued

            df.loc[ss_idx, "_ss_raw_value"] = raw_ss.values
            df.loc[ss_idx, "_ss_effective_value"] = effective_ss.values
            df.loc[ss_idx, "_ss_fallback_used"] = (raw_ss.le(0.0) & effective_ss.gt(0.0)).values
            df.loc[ss_idx, "_ss_zero_streak"] = zero_streak.values

        planned_cols = [date_map[d + timedelta(days=7 * off)] for d in planned_weeks if d + timedelta(days=7 * off) in date_map]

        # Firm PO for SI at the candidate balancing week is cumulative from the selected
        # Firm PO Week through the candidate Balance Week. This captures any additional
        # Firm already present in the Timeline during the weeks between Firm PO Week and
        # Balance Week.
        firm_range_cols = [date_map[d + timedelta(days=7 * off)] for d in firm_eval_weeks if d + timedelta(days=7 * off) in date_map]
        if firm_range_cols:
            df.loc[idx, "_firm_through_balance_value"] = df.loc[idx, firm_range_cols].apply(
                pd.to_numeric, errors="coerce"
            ).fillna(0.0).sum(axis=1).values

        # Keep the Firm PO at the selected Firm PO Week separately for audit / reconciliation.
        firm_src = firm_week + timedelta(days=7 * off)
        firm_col = date_map.get(firm_src)
        if firm_col:
            df.loc[idx, "_firm_week_value"] = pd.to_numeric(df.loc[idx, firm_col], errors="coerce").fillna(0.0).values
        if planned_cols:
            df.loc[idx, "_planned_sum"] = df.loc[idx, planned_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0).sum(axis=1).values
        # NET FCST is a demand/coverage calculation from the actual Current Week
        # through the balancing week. It is NOT shifted by warehouse ETA->ETD offset.
        # This is especially important for Whse 335.
        if net_weeks:
            direct_net_cols = [date_map[d] for d in net_weeks if d in date_map]
            if direct_net_cols:
                idx_335 = df.loc[idx, "Whse"].astype(str).eq("335")
                if idx_335.any():
                    rows_335 = df.loc[idx].loc[idx_335]
                    df.loc[rows_335.index, "_net_sum"] = rows_335[direct_net_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0).sum(axis=1).values
                idx_other = ~idx_335
                if idx_other.any():
                    rows_other = df.loc[idx].loc[idx_other]
                    net_cols = [date_map[d + timedelta(days=7 * off)] for d in net_weeks if d + timedelta(days=7 * off) in date_map]
                    if net_cols:
                        df.loc[rows_other.index, "_net_sum"] = rows_other[net_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0).sum(axis=1).values

    keys = ["Item #", "Whse", "Coll. Class"]
    def agg(dtype, col, name):
        part = df[df["Data Type"] == dtype].copy()
        return group_value(part, keys, col, name)
    si_g = agg("SHIPPABLE INV", "_target_value", "Base_SI")
    planned_g = agg("PLANNED POS", "_planned_sum", "PlannedPO_Sum")
    firm_g = agg("FIRM POS", "_firm_week_value", "FirmPO_FirmWeek")
    firm_through_g = agg("FIRM POS", "_firm_through_balance_value", "FirmPO_ThroughBalance")
    net_g = agg("NET FCST", "_net_sum", "NetFcst_Sum")
    # SS_Wk3 is the EFFECTIVE Safety Stock used by the optimizer. Raw_SS_Wk3
    # is retained internally for audit so candidate-week logic can tell when
    # Last Positive SS was used.
    ss_g = agg("SAFETY STK", "_ss_effective_value", "SS_Wk3")
    ss_raw_g = agg("SAFETY STK", "_ss_raw_value", "Raw_SS_Wk3")
    ss_fb = df[df["Data Type"] == "SAFETY STK"].groupby(keys, dropna=False)["_ss_fallback_used"].max().reset_index().rename(columns={"_ss_fallback_used": "SS_Fallback_Used"})
    ss_zero = df[df["Data Type"] == "SAFETY STK"].groupby(keys, dropna=False)["_ss_zero_streak"].max().reset_index().rename(columns={"_ss_zero_streak": "SS_Zero_Streak"})
    base = df[df["Data Type"].isin(["SHIPPABLE INV", "PLANNED POS", "FIRM POS", "NET FCST", "SAFETY STK"])][keys].drop_duplicates()
    out = (
        base.merge(si_g, on=keys, how="left")
            .merge(planned_g, on=keys, how="left")
            .merge(firm_g, on=keys, how="left")
            .merge(firm_through_g, on=keys, how="left")
            .merge(net_g, on=keys, how="left")
            .merge(ss_g, on=keys, how="left")
            .merge(ss_raw_g, on=keys, how="left")
            .merge(ss_fb, on=keys, how="left")
            .merge(ss_zero, on=keys, how="left")
    )
    for c in ["Base_SI", "PlannedPO_Sum", "FirmPO_FirmWeek", "FirmPO_ThroughBalance", "NetFcst_Sum", "SS_Wk3", "Raw_SS_Wk3"]:
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0.0)
    out["SS_Fallback_Used"] = out.get("SS_Fallback_Used", False).fillna(False).astype(bool)
    out["SS_Zero_Streak"] = pd.to_numeric(out.get("SS_Zero_Streak", 0), errors="coerce").fillna(0).astype(int)
    # Audit aliases: Firm PO at Firm PO Week and cumulative Firm PO through Balance Week.
    out["Timeline Firm PO"] = out["FirmPO_FirmWeek"]
    out["Timeline Firm PO Through Balance"] = out["FirmPO_ThroughBalance"]
    # Auto Separate uses a single, explicit SI basis:
    #   SI Before Supply @ Balance Week
    #   = Base SI @ Balance Week
    #     - Planned POS through Balance Week
    #     - ALL Firm PO from Firm PO Week through Balance Week
    # Firm PO at the selected Firm PO Week remains stored separately for reconciliation,
    # while the SI baseline uses the cumulative Firm PO through the candidate Balance Week.
    out["Current SI"] = out["Base_SI"] - out["PlannedPO_Sum"] - out["FirmPO_ThroughBalance"]
    mask_335 = out["Whse"].astype(str) == "335"
    out.loc[mask_335, "Current SI"] = (
        out.loc[mask_335, "Base_SI"]
        - out.loc[mask_335, "PlannedPO_Sum"]
        - out.loc[mask_335, "FirmPO_ThroughBalance"]
        + out.loc[mask_335, "NetFcst_Sum"]
    )
    out["FirmPO_Target"] = out["FirmPO_FirmWeek"]
    out["Current SI-SS"] = out["Current SI"] - out["SS_Wk3"]
    out["Current SS%"] = out.apply(lambda r: safe_ss_ratio(float(r["Current SI"]), float(r["SS_Wk3"])), axis=1)
    out["Balance Week"] = eval_week
    vendor_col = next((c for c in df.columns if c.lower() == "vendor"), None)
    if vendor_col:
        vendor_map = df.groupby(keys, dropna=False)[vendor_col].first().reset_index().rename(columns={vendor_col: "Vendor"})
        out = out.merge(vendor_map, on=keys, how="left")
    else:
        out["Vendor"] = ""
    out = out.rename(columns={"Item #": "Item", "Coll. Class": "ProdResourceID"})
    out["Item"] = out["Item"].map(normalize_item)
    out["Whse"] = out["Whse"].map(normalize_whse)
    debug = pd.DataFrame([["Evaluation week", fmt_date(eval_week)], ["Rows", len(out)]], columns=["Field", "Value"])
    offset_by_whse = df[["Whse", "_offset_weeks"]].drop_duplicates().rename(columns={"_offset_weeks": "Used Offset Weeks"}).sort_values("Whse")
    return out, debug, offset_by_whse


def _build_balance_candidate_input(
    raw: pd.DataFrame,
    date_map: Dict[date, str],
    first_etd_week: date,
    balance_week: date,
    current_week: date,
    offset_map: Dict[str, int],
    fixed_supply: pd.DataFrame,
    firm_week: Optional[date] = None,
) -> pd.DataFrame:
    """Build a fresh optimizer input for one balancing week while keeping Firm PO tied to firm_week.

    Inventory metrics are recalculated from the raw PlanDetailTimeline at the candidate balancing
    week. The SI baseline subtracts all Firm PO from the selected Firm PO Week through the candidate
    Balance Week, inclusive. Main/other vendor supply and reconciliation values are then applied once.
    """
    metrics, _, _ = compute_plan_metrics_for_week_from_raw(
        raw, date_map, first_etd_week, balance_week, current_week, offset_map,
        firm_week=firm_week or balance_week,
    )
    if metrics.empty:
        return pd.DataFrame()

    supply_cols = [
        "Item", "Whse", "Main Vendor", "Main Vendor F Wk3", "Other Vendor Supply",
        "Other Vendor List", "Timeline Firm PO", "PSW F Used for Reconciliation",
        "Firm PO Reconciliation Gap", "Total Supply Added to SI"
    ] + [c for c in PLAN_METADATA_COLUMNS if c in fixed_supply.columns]
    available = [c for c in supply_cols if c in fixed_supply.columns]
    out = metrics.merge(fixed_supply[available].drop_duplicates(), on=["Item", "Whse"], how="left")

    for c in [
        "Main Vendor F Wk3", "Other Vendor Supply", "Timeline Firm PO",
        "PSW F Used for Reconciliation", "Firm PO Reconciliation Gap", "Total Supply Added to SI"
    ]:
        if c not in out.columns:
            out[c] = 0.0
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0.0)
    if "Other Vendor List" not in out.columns:
        out["Other Vendor List"] = ""

    # Firm PO remains tied to the original firm week. Only the inventory evaluation point changes.
    out["F Wk3"] = out["Main Vendor F Wk3"]
    out["PSW F Used for Reconciliation"] = out["Main Vendor F Wk3"] + out["Other Vendor Supply"]
    out["Total Supply Added to SI"] = (
        out["Main Vendor F Wk3"]
        + out["Other Vendor Supply"]
        + out["Firm PO Reconciliation Gap"]
    )
    out["Original SI Before Supply"] = out["Current SI"]
    out["Original SI-SS Before Supply"] = out["Current SI-SS"]
    out["New SI"] = out["Current SI"] + out["Total Supply Added to SI"]
    out["New SI-SS"] = out["Current SI-SS"] + out["Total Supply Added to SI"]
    out["Sum of SI Wk3"] = out["New SI"]
    out["Sum of SI-SS Wk3"] = out["New SI-SS"]
    out["Current SI"] = out["New SI"]
    out["Current SI-SS"] = out["New SI-SS"]
    out["Average of SS Wk3"] = pd.to_numeric(out["SS_Wk3"], errors="coerce").fillna(0.0)
    out["Raw SS Wk3"] = pd.to_numeric(out.get("Raw_SS_Wk3", out["SS_Wk3"]), errors="coerce").fillna(0.0)
    out["Last Positive SS Used"] = out.get("SS_Fallback_Used", False).fillna(False).astype(bool)
    out["SS Zero Streak"] = pd.to_numeric(out.get("SS_Zero_Streak", 0), errors="coerce").fillna(0).astype(int)
    out["Current SS%"] = out.apply(
        lambda r: safe_ss_ratio(float(r["Current SI"]), float(r["Average of SS Wk3"])), axis=1
    )
    out["Firm PO Total"] = out.groupby("Item")["F Wk3"].transform("sum")
    out["Balance Week"] = balance_week
    # Important: SI/supply has already been calculated here. prepare_optimizer_input()
    # must not add Total Supply Added to SI again.
    out["_supply_already_applied"] = True
    return out


def _allocate_sub_balance_candidate(
    main_allocated: pd.DataFrame,
    item_rules: Dict[str, PriorityRule],
    respect_priority_rank: bool = False,
    freeze_hold_buy: bool = False,
) -> pd.DataFrame:
    """Allocate sub-vendor Firm PO for one item/candidate week using Main Vendor SI After as baseline."""
    g = main_allocated.copy().reset_index(drop=True)
    orig = pd.to_numeric(g.get("Other Vendor Supply", 0.0), errors="coerce").fillna(0.0).round().astype(int)
    temp = g.copy()
    temp["F Wk3"] = orig
    temp["Current SI"] = pd.to_numeric(temp["Current SI After"], errors="coerce").fillna(0.0)
    temp["Average of SS Wk3"] = pd.to_numeric(temp["Average of SS Wk3"], errors="coerce").fillna(0.0)
    temp["Firm PO Total"] = int(orig.sum())
    if "Sub Future Firm Floor" in g.columns:
        temp["Future Firm Floor"] = pd.to_numeric(g["Sub Future Firm Floor"], errors="coerce").fillna(0.0)

    if int(orig.sum()) == 0:
        final = orig.copy()
        # Reuse Main hard-lock status for audit/reference eligibility. The same SI/HB rules apply.
        sub_hard_lock = g.get("Hard Lock", pd.Series(False, index=g.index)).fillna(False).astype(bool)
    else:
        allocated = allocate_item(
            temp,
            item_rules,
            respect_priority_rank=respect_priority_rank,
            freeze_hold_buy=freeze_hold_buy,
        )
        final = pd.to_numeric(allocated["F Wk3 After Destination Change"], errors="coerce").fillna(0).astype(int)
        sub_hard_lock = allocated.get("Hard Lock", pd.Series(False, index=allocated.index)).fillna(False).astype(bool)

    out = g.copy()
    out["Sub Vendor F Original"] = orig.to_numpy()
    out["Sub Vendor F After Destination Change"] = final.to_numpy()
    out["Sub Vendor Net Destination Change"] = out["Sub Vendor F After Destination Change"] - out["Sub Vendor F Original"]
    out["Sub Vendor SI Before"] = pd.to_numeric(out["Current SI After"], errors="coerce").fillna(0.0)
    out["Sub Vendor SI After"] = out["Sub Vendor SI Before"] + out["Sub Vendor Net Destination Change"]
    out["Sub Vendor SS% After"] = out.apply(
        lambda r: safe_ss_ratio(float(r["Sub Vendor SI After"]), float(r["Average of SS Wk3"])),
        axis=1,
    )
    out["Sub Hard Lock"] = list(sub_hard_lock)
    return out


def _evaluate_active_balance_health(
    allocated: pd.DataFrame,
    after_col: str,
    ratio_col: str,
    threshold: float,
    future_min_si_by_whse: Optional[Dict[str, float]] = None,
    hard_lock_col: str = "Hard Lock",
) -> dict:
    """Evaluate one candidate week using the approved Hybrid Balance rule.

    Priority of checks:
    1) Normal planning: Active WH (Firm After > 0) with meaningful SS must be <= Healthy SS%.
    2) Future shortage protection: any change-eligible Buy WH must stay at/above SI=0 through the
       remaining fully-evaluable horizon after applying this candidate's destination-change delta.
       This catches demand peaks even when the warehouse is Neutral / Firm After = 0 today.
    3) Runout balance: when raw SS has been zero for 2+ consecutive candidate buckets (or there is
       no positive SS history), SS% no longer gates the week; future SI >= 0 is the governing test.

    Firm After = 0 remains excluded from the Healthy SS% gate for this candidate only.
    It can still be protected by the separate future-shortage guard if it is change-eligible.
    """
    future_min_si_by_whse = future_min_si_by_whse or {}
    empty = {
        "healthy": False,
        "active_whs": [],
        "excluded_zero_whs": [],
        "above_whs": [],
        "active_count": 0,
        "excluded_zero_count": 0,
        "above_count": 0,
        "min_active_ratio": None,
        "max_active_ratio": None,
        "last_positive_ss_whs": [],
        "no_positive_ss_whs": [],
        "runout_whs": [],
        "future_shortage_whs": [],
        "min_future_si": None,
    }
    if allocated is None or allocated.empty:
        return empty

    x = allocated.copy()
    x["_whse"] = x["Whse"].map(normalize_whse)
    x["_after"] = pd.to_numeric(x.get(after_col, 0.0), errors="coerce").fillna(0.0)
    x["_ratio"] = pd.to_numeric(x.get(ratio_col, float("nan")), errors="coerce")
    x["_raw_ss"] = pd.to_numeric(x.get("Raw SS Wk3", x.get("Average of SS Wk3", 0.0)), errors="coerce").fillna(0.0)
    x["_effective_ss"] = pd.to_numeric(x.get("Average of SS Wk3", 0.0), errors="coerce").fillna(0.0)
    x["_zero_streak"] = pd.to_numeric(x.get("SS Zero Streak", 0), errors="coerce").fillna(0).astype(int)
    x["_last_positive_used"] = x.get("Last Positive SS Used", pd.Series(False, index=x.index)).fillna(False).astype(bool)
    x["_hard"] = x.get(hard_lock_col, x.get("Hard Lock", pd.Series(False, index=x.index))).fillna(False).astype(bool)
    x["_firm_zero"] = x.get("Firm Zero Target", pd.Series(False, index=x.index)).fillna(False).astype(bool)
    x["_future_min_si"] = x["_whse"].map(lambda w: future_min_si_by_whse.get(str(w), float("nan")))

    active = x[x["_after"] > 0].copy()
    excluded = x[x["_after"] <= 0].copy()

    active_whs = list(dict.fromkeys(active["_whse"].astype(str).tolist()))
    excluded_whs = list(dict.fromkeys(excluded["_whse"].astype(str).tolist()))
    last_positive_whs = list(dict.fromkeys(x.loc[x["_last_positive_used"], "_whse"].astype(str).tolist()))
    no_positive_mask = (x["_raw_ss"] <= 0) & (x["_effective_ss"] <= 0)
    no_positive_whs = list(dict.fromkeys(x.loc[no_positive_mask, "_whse"].astype(str).tolist()))

    # SS=0 on the first candidate bucket uses Last Positive SS as a bridge.
    # From the second consecutive zero bucket onward, SS% no longer has business meaning.
    runout_mask = (x["_zero_streak"] >= 2) | no_positive_mask
    runout_whs = list(dict.fromkeys(x.loc[runout_mask, "_whse"].astype(str).tolist()))

    normal_active = active[~runout_mask.loc[active.index]].copy()
    ratios = pd.to_numeric(normal_active["_ratio"], errors="coerce")
    pass_mask = ratios.notna() & ratios.le(float(threshold))
    above_whs = list(dict.fromkeys(normal_active.loc[~pass_mask, "_whse"].astype(str).tolist()))
    finite_vals = [float(v) for v in pd.to_numeric(active["_ratio"], errors="coerce").tolist() if pd.notna(v) and math.isfinite(float(v))]

    # Forward guard covers ALL change-eligible B warehouses, including Neutral/Firm=0-after rows,
    # so a future demand peak can prevent a too-early Balance Week. User hard locks and FIRM=0
    # targets are excluded because the optimizer is not allowed to repair them.
    guard = x[(~x["_hard"]) & (~x["_firm_zero"])].copy()
    guard_known = guard[guard["_future_min_si"].notna()].copy()
    shortage = guard_known[guard_known["_future_min_si"] < -1e-9]
    future_shortage_whs = list(dict.fromkeys(shortage["_whse"].astype(str).tolist()))
    min_future_si = None
    if not guard_known.empty:
        vals = pd.to_numeric(guard_known["_future_min_si"], errors="coerce").dropna().tolist()
        if vals:
            min_future_si = float(min(vals))

    # Runout Active rows pass their SS gate only when future SI is safe. This is normally already
    # enforced by the all-B forward guard, but keep the explicit test for transparent audit behavior.
    runout_active = active[runout_mask.loc[active.index]].copy()
    runout_fail_whs = []
    for wh in runout_active["_whse"].astype(str).tolist():
        v = future_min_si_by_whse.get(wh)
        if v is None or (pd.notna(v) and float(v) < -1e-9):
            runout_fail_whs.append(wh)
    runout_fail_whs = list(dict.fromkeys(runout_fail_whs))

    healthy = bool(pass_mask.all()) and not future_shortage_whs and not runout_fail_whs
    return {
        "healthy": healthy,
        "active_whs": active_whs,
        "excluded_zero_whs": excluded_whs,
        "above_whs": above_whs,
        "active_count": len(active_whs),
        "excluded_zero_count": len(excluded_whs),
        "above_count": len(above_whs),
        "min_active_ratio": min(finite_vals) if finite_vals else None,
        "max_active_ratio": max(finite_vals) if finite_vals else None,
        "last_positive_ss_whs": last_positive_whs,
        "no_positive_ss_whs": no_positive_whs,
        "runout_whs": runout_whs,
        "future_shortage_whs": future_shortage_whs,
        "min_future_si": min_future_si,
    }


def _get_balance_candidate_cached(
    cache: Dict[date, pd.DataFrame],
    raw: pd.DataFrame,
    date_map: Dict[date, str],
    first_etd_week: date,
    wk: date,
    current_week: date,
    offset_map: Dict[str, int],
    fixed_supply: pd.DataFrame,
    firm_week: date,
) -> pd.DataFrame:
    if wk not in cache:
        cache[wk] = _build_balance_candidate_input(
            raw, date_map, first_etd_week, wk, current_week, offset_map, fixed_supply, firm_week=firm_week
        )
    return cache[wk]


def _future_min_si_original_map(
    item: str,
    start_week: date,
    end_week: date,
    cache: Dict[date, pd.DataFrame],
    raw: pd.DataFrame,
    date_map: Dict[date, str],
    first_etd_week: date,
    current_week: date,
    offset_map: Dict[str, int],
    fixed_supply: pd.DataFrame,
    firm_week: date,
) -> Dict[str, float]:
    """Minimum projected SI by warehouse under the ORIGINAL Firm distribution.

    Candidate inputs already include the original main/sub supply and reconciliation. A destination-change
    delta is persistent across later weeks, so future SI after a candidate allocation can be obtained as:
        future_min_after = future_min_original + Net Destination Change.
    """
    mins: Dict[str, float] = {}
    for fw in date_range_saturdays(start_week, end_week):
        cand = _get_balance_candidate_cached(
            cache, raw, date_map, first_etd_week, fw, current_week, offset_map, fixed_supply, firm_week
        )
        if cand is None or cand.empty:
            continue
        g = cand[cand["Item"].astype(str) == str(item)].copy()
        if g.empty:
            continue
        for _, r in g.iterrows():
            wh = normalize_whse(r.get("Whse", ""))
            if not wh:
                continue
            v = pd.to_numeric(pd.Series([r.get("Current SI", float("nan"))]), errors="coerce").iloc[0]
            if pd.isna(v):
                continue
            fv = float(v)
            mins[wh] = min(mins.get(wh, fv), fv)
    return mins


def _attach_future_firm_floor(
    group: pd.DataFrame,
    driver: str,
    future_min_original: Dict[str, float],
) -> pd.DataFrame:
    """Attach minimum Firm After needed to keep future SI >= 0.

    If future_min_original is measured with original Firm distribution, moving from OrigFirm to FinalFirm
    shifts every projected future SI by (FinalFirm - OrigFirm). Therefore the minimum safe FinalFirm is:
        ceil(OrigFirm - future_min_original), floored at zero.
    """
    g = group.copy()
    orig_col = "F Wk3" if str(driver).upper() == "MAIN" else "Other Vendor Supply"
    floors = []
    min_vals = []
    for _, r in g.iterrows():
        wh = normalize_whse(r.get("Whse", ""))
        min_si = future_min_original.get(wh)
        min_vals.append(min_si if min_si is not None else float("nan"))
        orig = float(pd.to_numeric(pd.Series([r.get(orig_col, 0.0)]), errors="coerce").fillna(0.0).iloc[0])
        if min_si is None or not math.isfinite(float(min_si)):
            floor = 0
        else:
            floor = max(0, int(math.ceil(orig - float(min_si) - 1e-12)))
        floors.append(floor)
    g["Future Min SI Original"] = min_vals
    if str(driver).upper() == "MAIN":
        g["Future Firm Floor"] = floors
    else:
        g["Sub Future Firm Floor"] = floors
    return g


def _future_min_after_map(
    evaluated: pd.DataFrame,
    driver: str,
    future_min_original: Dict[str, float],
) -> Dict[str, float]:
    delta_col = "Net Destination Change" if str(driver).upper() == "MAIN" else "Sub Vendor Net Destination Change"
    out: Dict[str, float] = {}
    for _, r in evaluated.iterrows():
        wh = normalize_whse(r.get("Whse", ""))
        if wh not in future_min_original:
            continue
        delta = float(pd.to_numeric(pd.Series([r.get(delta_col, 0.0)]), errors="coerce").fillna(0.0).iloc[0])
        out[wh] = float(future_min_original[wh]) + delta
    return out


def _dynamic_balance_candidate_weeks(
    date_map: Dict[date, str],
    firm_week: date,
    offset_map: Dict[str, int],
    whses: Optional[Iterable[str]] = None,
) -> List[date]:
    """Return candidate ETD weeks through the last fully evaluable PlanDetail week.

    PlanDetail columns are ETA-side weekly buckets. Candidate ETD weeks are shifted
    forward by warehouse transit offsets when read. To avoid inventing zero values
    beyond the source horizon, the last candidate is capped at the latest raw week
    minus the largest positive offset among the warehouses relevant to this item.
    There is intentionally no fixed Week-14 limit.
    """
    if not date_map:
        return [firm_week]
    max_raw_week = max(date_map.keys())
    relevant_whses = [normalize_whse(w) for w in whses] if whses is not None else None
    values = []
    if relevant_whses is None:
        values = list((offset_map or {}).values())
    else:
        values = [(offset_map or {}).get(w, 0) for w in relevant_whses]
    offsets = []
    for v in values:
        try:
            offsets.append(max(0, int(v)))
        except Exception:
            pass
    max_offset = max(offsets) if offsets else 0
    last_eval = max_raw_week - timedelta(days=7 * max_offset)
    if last_eval < firm_week:
        return [firm_week]
    return date_range_saturdays(firm_week, last_eval)


def auto_select_balance_week_map(
    raw: pd.DataFrame,
    date_map: Dict[date, str],
    first_etd_week: date,
    firm_week: date,
    current_week: date,
    offset_map: Dict[str, int],
    fixed_supply: pd.DataFrame,
    priority_rules: Optional[Dict[str, PriorityRule]] = None,
    respect_priority_rank: bool = False,
    threshold: float = 1.5,
    freeze_hold_buy: bool = False,
) -> Tuple[Dict[str, date], pd.DataFrame]:
    """Select the earliest Balance Week using Hybrid Active-Firm / Future-Need / Runout logic.

    Approved rules per item:
    - Fixed user Healthy SS%; no Reference / Adjusted Healthy.
    - Evaluate candidate weeks sequentially from Firm PO Week through the last fully evaluable PlanDetail week.
    - Rebuild every candidate fresh and rerun Destination Change across all Buy (B) warehouses.
    - Driver: Main when Main Firm > 0; otherwise Sub when Main Firm = 0 and Sub Firm > 0.
    - Active WH = driver Firm After > 0. Firm After = 0 is excluded from the Healthy SS% gate for that
      candidate only and can become Active again next week.
    - Normal planning: Active WH with meaningful SS must be <= Healthy SS%.
    - Future shortage protection: all change-eligible B warehouses are checked through the remaining
      fully evaluable horizon. A demand peak that would drive projected SI below zero rejects the candidate.
      The allocator first protects a minimum Future Firm Floor before ordinary SS balancing.
    - Last Positive SS bridges the FIRST zero-SS bucket. From 2 consecutive zero-SS buckets onward,
      Runout Balance Mode replaces SS% with future SI >= 0 as the health condition.
    - If no candidate passes before the source horizon ends, use the last fully evaluable week as fallback.
    """
    priority_rules = priority_rules or {}
    items = raw["Item #"].map(normalize_item).dropna().astype(str).unique().tolist()

    horizon_src = raw.copy()
    horizon_src["_item"] = horizon_src["Item #"].map(normalize_item)
    horizon_src["_whse"] = horizon_src["Whse"].map(normalize_whse)
    if "MakeBuy Code" in horizon_src.columns:
        horizon_src = horizon_src[horizon_src["MakeBuy Code"].fillna("").astype(str).str.strip().str.upper().eq("B")]
    item_last_weeks: Dict[str, date] = {}
    for item in items:
        whses = horizon_src.loc[horizon_src["_item"].astype(str).eq(str(item)), "_whse"].dropna().astype(str).unique().tolist()
        item_weeks = _dynamic_balance_candidate_weeks(date_map, firm_week, offset_map, whses=whses)
        item_last_weeks[item] = item_weeks[-1]
    global_last_week = max(item_last_weeks.values()) if item_last_weeks else firm_week
    all_weeks = date_range_saturdays(firm_week, global_last_week)

    selected: Dict[str, date] = {}
    drivers: Dict[str, str] = {}
    debug_rows = []
    candidate_cache: Dict[date, pd.DataFrame] = {}

    wk1_input = _get_balance_candidate_cached(
        candidate_cache, raw, date_map, first_etd_week, firm_week, current_week, offset_map, fixed_supply, firm_week
    )

    def _debug_base(item, driver, wk, health, reason):
        return {
            "Item": item,
            "Driver Vendor": driver,
            "Selected Balance Week": fmt_date(wk),
            "Healthy SS%": threshold,
            "Active WHs": ", ".join(health.get("active_whs", [])),
            "Active WH Count": health.get("active_count"),
            "Firm=0 Excluded WHs": ", ".join(health.get("excluded_zero_whs", [])),
            "Firm=0 Excluded Count": health.get("excluded_zero_count"),
            "Active Above Healthy WHs": ", ".join(health.get("above_whs", [])),
            "Active Above Healthy Count": health.get("above_count"),
            "Min Active SS% After": health.get("min_active_ratio"),
            "Max Active SS% After": health.get("max_active_ratio"),
            "Last Positive SS WHs": ", ".join(health.get("last_positive_ss_whs", [])),
            "Runout WHs": ", ".join(health.get("runout_whs", [])),
            "Future Shortage WHs": ", ".join(health.get("future_shortage_whs", [])),
            "Min Future SI": health.get("min_future_si"),
            "No Positive SS History WHs": ", ".join(health.get("no_positive_ss_whs", [])),
            "Reason": reason,
        }

    for item in items:
        g = wk1_input[wk1_input["Item"].astype(str) == str(item)].copy() if not wk1_input.empty else pd.DataFrame()
        if g.empty:
            drivers[item] = "NONE"
            selected[item] = firm_week
            debug_rows.append(_debug_base(item, "NONE", firm_week, {}, "No Firm-week candidate rows; keep Firm PO Week"))
            continue
        main_pool = int(round(pd.to_numeric(g.get("F Wk3", 0.0), errors="coerce").fillna(0.0).sum()))
        sub_pool = int(round(pd.to_numeric(g.get("Other Vendor Supply", 0.0), errors="coerce").fillna(0.0).sum()))
        if main_pool > 0:
            drivers[item] = "MAIN"
        elif sub_pool > 0:
            drivers[item] = "SUB"
        else:
            drivers[item] = "NONE"
            selected[item] = firm_week
            debug_rows.append(_debug_base(item, "NONE", firm_week, {}, "No Main/Sub Firm PO at Firm PO Week; keep Firm PO Week"))

    unresolved = [item for item in items if item not in selected]

    for wk in all_weeks:
        if not unresolved:
            break
        candidate = _get_balance_candidate_cached(
            candidate_cache, raw, date_map, first_etd_week, wk, current_week, offset_map, fixed_supply, firm_week
        )
        if candidate.empty:
            continue

        still_unresolved = []
        for item in unresolved:
            driver = drivers.get(item, "NONE")
            g = candidate[candidate["Item"].astype(str) == str(item)].copy()
            if g.empty:
                still_unresolved.append(item)
                continue

            last_week = item_last_weeks.get(item, all_weeks[-1])
            future_min_original = _future_min_si_original_map(
                item, wk, last_week, candidate_cache, raw, date_map, first_etd_week,
                current_week, offset_map, fixed_supply, firm_week,
            )
            g = _attach_future_firm_floor(g, driver, future_min_original)

            item_whse = set(g["Whse"].astype(str).tolist())
            item_rules = {w: r for w, r in priority_rules.items() if w in item_whse}
            main_alloc = allocate_item(
                g.copy(), item_rules,
                respect_priority_rank=respect_priority_rank,
                freeze_hold_buy=freeze_hold_buy,
            )

            if driver == "MAIN":
                evaluated = main_alloc
                future_after = _future_min_after_map(evaluated, "MAIN", future_min_original)
                health = _evaluate_active_balance_health(
                    evaluated,
                    "F Wk3 After Destination Change",
                    "SS % After",
                    threshold,
                    future_min_si_by_whse=future_after,
                    hard_lock_col="Hard Lock",
                )
            elif driver == "SUB":
                evaluated = _allocate_sub_balance_candidate(
                    main_alloc,
                    item_rules,
                    respect_priority_rank=respect_priority_rank,
                    freeze_hold_buy=freeze_hold_buy,
                )
                future_after = _future_min_after_map(evaluated, "SUB", future_min_original)
                health = _evaluate_active_balance_health(
                    evaluated,
                    "Sub Vendor F After Destination Change",
                    "Sub Vendor SS% After",
                    threshold,
                    future_min_si_by_whse=future_after,
                    hard_lock_col="Sub Hard Lock",
                )
            else:
                selected[item] = wk
                continue

            if health["healthy"] or wk >= last_week:
                selected[item] = wk
                if health["healthy"]:
                    parts = []
                    if health.get("runout_whs"):
                        parts.append("Runout mode used for SS=0 continuation")
                    if not health.get("future_shortage_whs"):
                        parts.append("future SI protected")
                    prefix = "Firm Week" if wk == firm_week else "Moved forward"
                    reason = f"{prefix}: all Active normal-planning WH are within Healthy SS%; " + "; ".join(parts)
                else:
                    reason = (
                        f"PlanDetail horizon reached; {health['above_count']} Active normal-planning WH remain above Healthy SS%"
                        f"; future shortage WHs={len(health.get('future_shortage_whs', []))}"
                    )
                debug_rows.append(_debug_base(item, driver, wk, health, reason))
            else:
                still_unresolved.append(item)

        unresolved = still_unresolved

    for item in unresolved:
        fallback = item_last_weeks.get(item, all_weeks[-1])
        selected[item] = fallback
        debug_rows.append(_debug_base(item, drivers.get(item, "NONE"), fallback, {}, "PlanDetail horizon fallback: candidate rows unavailable"))

    debug_cols = [
        "Item", "Driver Vendor", "Selected Balance Week", "Healthy SS%",
        "Active WHs", "Active WH Count", "Firm=0 Excluded WHs", "Firm=0 Excluded Count",
        "Active Above Healthy WHs", "Active Above Healthy Count",
        "Min Active SS% After", "Max Active SS% After",
        "Last Positive SS WHs", "Runout WHs", "Future Shortage WHs", "Min Future SI",
        "No Positive SS History WHs", "Reason",
    ]
    return selected, pd.DataFrame(debug_rows, columns=debug_cols)


def build_optimizer_input_auto_balance(
    plan_csv_path: str,
    offset_map: Dict[str, int],
    firm_week: date,
    current_week: date,
    psw_supply_detail: pd.DataFrame,
    vendor_offset_maps: Dict[str, Dict[str, int]],
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    raw, date_map, first_original_week, last_original_week, first_etd_week = load_plan_source(plan_csv_path)
    selected_map, balance_debug = auto_select_balance_week_map(raw, date_map, first_etd_week, firm_week, current_week, offset_map)

    # Build firm-week supply once. This preserves the firm's selected week while metrics are evaluated at the chosen balance week.
    firm_fallback = pd.DataFrame(columns=["Item", "Whse", "F Wk3", "Main Vendor", "Main Vendor F Wk3", "Other Vendor Supply", "Other Vendor List", "Timeline Firm PO", "PSW F Used for Reconciliation", "Firm PO Reconciliation Gap", "Total Supply Added to SI"])
    if psw_supply_detail is not None and not psw_supply_detail.empty:
        supply_grouped, _, _ = split_main_other_vendor_supply(
            raw[["Item #", "Whse"]].drop_duplicates().rename(columns={"Item #": "Item"}).assign(Vendor=""),
            psw_supply_detail, firm_week, current_week, vendor_offset_maps=vendor_offset_maps,
        )
        # supply_grouped is generated against Timeline vendor if available; rebuild base rows with real Vendor below.
        try:
            raw_vendor_col = find_vendor_col(raw)
            if raw_vendor_col:
                timeline_vendor = raw[["Item #", "Whse", "Coll. Class", raw_vendor_col]].copy()
                timeline_vendor = timeline_vendor.rename(columns={"Item #":"Item", raw_vendor_col:"Vendor"})
                timeline_vendor["Item"] = timeline_vendor["Item"].map(normalize_item)
                timeline_vendor["Whse"] = timeline_vendor["Whse"].map(normalize_whse)
                supply_grouped, _, _ = split_main_other_vendor_supply(
                    timeline_vendor[["Item", "Whse", "Vendor"]].drop_duplicates(), psw_supply_detail, firm_week, current_week,
                    vendor_offset_maps=vendor_offset_maps,
                )
        except Exception:
            pass
        firm_fallback = supply_grouped.copy()
    if firm_fallback.empty:
        # No explicit PSW detail: use Production Schedule main quantity from the firm week.
        temp_prod, _ = load_fwk3_from_production(psw_supply_detail.attrs.get("__production_path", "") if psw_supply_detail is not None else "", firm_week) if False else (pd.DataFrame(), None)

    frames = []
    plan_offset_rows = []
    build_rows = []
    unique_weeks = sorted(set(selected_map.values()))
    for eval_week in unique_weeks:
        metrics, _, offset_dbg = compute_plan_metrics_for_week_from_raw(
            raw, date_map, first_etd_week, eval_week, current_week, offset_map, firm_week=firm_week
        )
        metrics["Item"] = metrics["Item"].astype(str)
        items_for_week = [item for item, wk in selected_map.items() if wk == eval_week]
        sub = metrics[metrics["Item"].isin(items_for_week)].copy()
        if not firm_fallback.empty:
            supply_cols = [c for c in ["Item", "Whse", "Main Vendor", "Main Vendor F Wk3", "Other Vendor Supply", "Other Vendor List", "Timeline Firm PO", "PSW F Used for Reconciliation", "Firm PO Reconciliation Gap", "Total Supply Added to SI"] if c in firm_fallback.columns]
            sub = sub.merge(firm_fallback[supply_cols].drop_duplicates(), on=["Item", "Whse"], how="left")
        for c in ["Main Vendor F Wk3", "Other Vendor Supply", "Firm PO Reconciliation Gap", "PSW F Used for Reconciliation", "Timeline Firm PO", "Total Supply Added to SI"]:
            if c not in sub.columns:
                sub[c] = 0.0
            sub[c] = pd.to_numeric(sub[c], errors="coerce").fillna(0.0)
        if "Other Vendor List" not in sub.columns:
            sub["Other Vendor List"] = ""
        sub["F Wk3"] = sub["Main Vendor F Wk3"]
        sub["Total Supply Added to SI"] = sub["Main Vendor F Wk3"] + sub["Other Vendor Supply"] + sub["Firm PO Reconciliation Gap"]
        sub["Original SI Before Supply"] = sub["Current SI"]
        sub["Original SI-SS Before Supply"] = sub["Current SI-SS"]
        sub["New SI"] = sub["Current SI"] + sub["Total Supply Added to SI"]
        sub["New SI-SS"] = sub["Current SI-SS"] + sub["Total Supply Added to SI"]
        sub["Sum of SI Wk3"] = sub["New SI"]
        sub["Sum of SI-SS Wk3"] = sub["New SI-SS"]
        sub["Current SI"] = sub["New SI"]
        sub["Current SI-SS"] = sub["New SI-SS"]
        sub["Current SS%"] = sub.apply(lambda r: safe_ss_ratio(float(r["Current SI"]), float(r["Average of SS Wk3"])), axis=1)
        sub["Firm PO Total"] = sub.groupby("Item")["F Wk3"].transform("sum")
        frames.append(sub)
        build_rows.append([fmt_date(eval_week), len(sub), float(sub["F Wk3"].sum()), len(items_for_week)])
        if not offset_dbg.empty:
            x = offset_dbg.copy(); x["Evaluation Week"] = fmt_date(eval_week); plan_offset_rows.extend(x.to_dict("records"))
    out = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    return out, pd.DataFrame(build_rows, columns=["Balance Week","Rows","Firm PO Total","Items"]), pd.DataFrame(columns=["Item","Whse","ProdResourceID"]), pd.DataFrame(plan_offset_rows), balance_debug


# ============================================================
# Step 5: Optimizer logic from destination_change_optimizer_phase2only.py
# ============================================================

@dataclass
class PriorityRule:
    whse: str
    # Supported modes: SI, SS, FIRM. FIRM currently means Firm PO target = 0.
    mode: str
    value: float
    rank: int = 9999


def normalize_pct(value) -> float:
    if value is None:
        return 0.0
    value = float(value)
    if value > 1 or value < -1:
        value = value / 100.0
    return value


def round_to_int_units(value: float) -> int:
    return int(round(float(value)))


def safe_ss_ratio(current_si: float, ss_target: float) -> float:
    if ss_target <= 0:
        if current_si > 0:
            return math.inf
        if current_si < 0:
            return -math.inf
        return 0.0
    return current_si / ss_target


def current_si_after(row: dict) -> int:
    return int(row["current_si"] + (row["final_f"] - row["orig_f"]))


def current_ss_after(row: dict) -> float:
    return safe_ss_ratio(current_si_after(row), row["ss_target"])


def compute_priority_target_final(row: dict, rule: PriorityRule) -> int:
    orig_f = row["orig_f"]
    current_si = row["current_si"]
    ss_target = row["ss_target"]
    if rule.mode == "SI":
        pct = max(0.0, min(1.0, float(rule.value)))
        target_si_after = current_si * (1.0 - pct)
        return max(0, round_to_int_units(orig_f + (target_si_after - current_si)))
    if rule.mode == "SS":
        target_si_after = ss_target * float(rule.value)
        return max(0, round_to_int_units(orig_f + (target_si_after - current_si)))
    if rule.mode == "FIRM":
        # Explicit Firm PO target mode. The current supported target is 0.
        return 0
    return int(orig_f)


def build_rows(group: pd.DataFrame, item_rules: Dict[str, PriorityRule], freeze_hold_buy: bool = False) -> Tuple[List[dict], int]:
    rows = []
    for _, r in group.iterrows():
        hold_buy = str(r.get("Hold/ Buy", "")).strip().upper() if freeze_hold_buy else ""
        row = {
            "item": r["Item"],
            "prod": r["ProdResourceID"],
            "whse": normalize_whse(r["Whse"]),
            "orig_f": round_to_int_units(r["F Wk3"]),
            "current_si": round_to_int_units(r["Current SI"]),
            "ss_target": float(r["Average of SS Wk3"]),
            "final_f": 0,
            "priority_rule_mode": "",
            "priority_rule_value": None,
            "priority_rank": 9999,
            "priority_target_f_after": None,
            "hard_lock": bool((freeze_hold_buy and hold_buy == "HB")),
            "hold_buy_frozen": bool((freeze_hold_buy and hold_buy == "HB")),
            "firm_zero_target": False,
            "exclude_from_recipient": False,
            # Auto-Separate-only forward-demand protection. Minimum absolute Firm After needed
            # to keep projected future SI >= 0 through the remaining evaluable horizon.
            "future_firm_floor": max(0, round_to_int_units(pd.to_numeric(pd.Series([r.get("Future Firm Floor", 0)]), errors="coerce").fillna(0.0).iloc[0])),
        }
        rule = item_rules.get(row["whse"])
        if rule is not None:
            row["priority_rule_mode"] = rule.mode
            row["priority_rule_value"] = rule.value
            row["priority_rank"] = int(getattr(rule, "rank", 9999) or 9999)
            # Priority SI = 0 is a hard lock: preserve the warehouse's original Firm PO quantity.
            priority_lock = bool(rule.mode == "SI" and max(0.0, min(1.0, float(rule.value))) <= 0.0)
            firm_zero = bool(rule.mode == "FIRM")
            row["hard_lock"] = bool(row["hard_lock"] or priority_lock)
            row["firm_zero_target"] = firm_zero
            row["exclude_from_recipient"] = firm_zero
            row["priority_target_f_after"] = compute_priority_target_final(row, rule)
        if row["hard_lock"]:
            row["final_f"] = row["orig_f"]
        rows.append(row)
    return rows, round_to_int_units(group["Firm PO Total"].iloc[0])


def choose_priority_recipient(rows: List[dict], priority_indices: List[int], respect_priority_rank: bool = False) -> Optional[int]:
    candidates = []
    for idx in priority_indices:
        row = rows[idx]
        target = row.get("priority_target_f_after")
        if target is None or row["final_f"] >= target or row.get("hard_lock") or row.get("exclude_from_recipient"):
            continue
        gap = target - row["final_f"]
        rank = int(row.get("priority_rank", 9999) or 9999)
        primary_metric = current_si_after(row) if row["priority_rule_mode"] == "SI" else current_ss_after(row)
        secondary_metric = current_ss_after(row) if row["priority_rule_mode"] == "SI" else current_si_after(row)
        candidates.append((rank, gap, primary_metric, secondary_metric, row["whse"], idx))
    if not candidates:
        return None
    if respect_priority_rank:
        min_rank = min(c[0] for c in candidates)
        candidates = [c for c in candidates if c[0] == min_rank]
    # Within a rank (or when rank is disabled): prioritize the largest remaining target gap,
    # then the existing rule metric, then warehouse code.
    candidates.sort(key=lambda x: (-x[1], x[2], x[3], x[4]))
    return candidates[0][5]


def choose_future_floor_recipient(rows: List[dict], candidate_indices: List[int]) -> Optional[int]:
    """Choose a non-priority recipient that is still below its future-safe Firm floor."""
    candidates = []
    for idx in candidate_indices:
        row = rows[idx]
        if row.get("hard_lock") or row.get("exclude_from_recipient"):
            continue
        floor = int(max(0, row.get("future_firm_floor", 0) or 0))
        gap = floor - int(row.get("final_f", 0))
        if gap <= 0:
            continue
        ratio = current_ss_after(row)
        candidates.append((-gap, ratio, row["whse"], idx))
    if not candidates:
        return None
    candidates.sort(key=lambda x: (x[0], x[1], x[2]))
    return candidates[0][3]


def choose_lowest_ss_recipient(rows: List[dict], candidate_indices: List[int]) -> Optional[int]:
    candidates = []
    for idx in candidate_indices:
        if rows[idx].get("hard_lock") or rows[idx].get("exclude_from_recipient"):
            continue
        after_si = current_si_after(rows[idx])
        ratio = safe_ss_ratio(after_si, rows[idx]["ss_target"])
        # Tie-breaker order: Lowest SS% After -> Highest SI After -> Warehouse code.
        candidates.append((ratio, -after_si, rows[idx]["whse"], idx))
    if not candidates:
        return None
    candidates.sort(key=lambda x: (x[0], x[1], x[2]))
    return candidates[0][3]


def allocate_item(group: pd.DataFrame, item_rules: Dict[str, PriorityRule], respect_priority_rank: bool = False, freeze_hold_buy: bool = False) -> pd.DataFrame:
    rows, total_f = build_rows(group, item_rules, freeze_hold_buy=freeze_hold_buy)
    locked_total = sum(int(r["orig_f"]) for r in rows if r.get("hard_lock"))
    remaining = total_f - locked_total
    if remaining < 0:
        raise ValueError(f"Item {group['Item'].iloc[0]}: hard-locked Firm PO exceeds total Firm PO.")

    priority_indices = [i for i, r in enumerate(rows) if r["priority_rule_mode"] and not r.get("hard_lock") and not r.get("exclude_from_recipient")]
    non_priority_indices = [i for i, r in enumerate(rows) if not r["priority_rule_mode"] and not r.get("hard_lock") and not r.get("exclude_from_recipient")]

    while remaining > 0:
        idx = choose_priority_recipient(rows, priority_indices, respect_priority_rank=respect_priority_rank)
        if idx is None:
            break
        rows[idx]["final_f"] += 1
        remaining -= 1

    allocation_pool = non_priority_indices[:] if non_priority_indices else priority_indices[:]

    # Before normal SS balancing, protect upcoming demand peaks for non-priority/change-eligible WHs.
    # Explicit user priority rules remain higher precedence; if they consume the available pool, the
    # forward-shortage guard will reject this candidate week instead of silently violating the rule.
    future_floor_pool = non_priority_indices[:] if non_priority_indices else []
    while remaining > 0 and future_floor_pool:
        idx = choose_future_floor_recipient(rows, future_floor_pool)
        if idx is None:
            break
        rows[idx]["final_f"] += 1
        remaining -= 1

    while remaining > 0 and allocation_pool:
        idx = choose_lowest_ss_recipient(rows, allocation_pool)
        if idx is None:
            break
        rows[idx]["final_f"] += 1
        remaining -= 1

    if remaining > 0:
        raise ValueError(f"Item {group['Item'].iloc[0]}: could not allocate all Firm PO quantity.")

    out = group.copy().reset_index(drop=True)
    out["F Wk3 Original"] = out["F Wk3"].round().astype(int)
    out["F Wk3 After Destination Change"] = [r["final_f"] for r in rows]
    out["Net Destination Change"] = out["F Wk3 After Destination Change"] - out["F Wk3 Original"]
    out["Current SI After"] = out["Current SI"] + out["Net Destination Change"]
    out["SS % After"] = out.apply(lambda r: safe_ss_ratio(float(r["Current SI After"]), float(r["Average of SS Wk3"])), axis=1)
    out["Remaining Unallocated PO"] = total_f - int(out["F Wk3 After Destination Change"].sum())
    out["Priority Rule Mode"] = [r["priority_rule_mode"] for r in rows]
    out["Priority Rule Value"] = [r["priority_rule_value"] for r in rows]
    out["Priority Rank"] = [r.get("priority_rank", 9999) for r in rows]
    out["Priority Target F After"] = [r["priority_target_f_after"] for r in rows]
    out["Hard Lock"] = [bool(r.get("hard_lock")) for r in rows]
    out["Firm Zero Target"] = [bool(r.get("firm_zero_target")) for r in rows]

    # Hard validation for FIRM=0. HB/SI hard locks take precedence.
    firm_zero_mask = out["Firm Zero Target"] & (~out["Hard Lock"])
    if firm_zero_mask.any():
        if not pd.to_numeric(out.loc[firm_zero_mask, "F Wk3 After Destination Change"], errors="coerce").fillna(0).eq(0).all():
            raise ValueError(f"Item {out['Item'].iloc[0]}: FIRM=0 priority was not enforced for all targeted warehouses.")

    if int(out["F Wk3 Original"].sum()) != int(out["F Wk3 After Destination Change"].sum()):
        raise ValueError(f"Item {out['Item'].iloc[0]}: Firm PO total is not preserved.")
    return out


def apply_zero_ss_equalization(detail_full: pd.DataFrame) -> pd.DataFrame:
    """Fallback pass for items with multiple zero-SS warehouses.

    Only non-priority, non-hard-locked warehouses with SS=0 participate. One Firm PO unit is
    moved from the highest-SI warehouse to the lowest-SI warehouse until the SI spread is <= 1
    or no further move is possible. Total Firm PO by item is preserved.
    """
    if detail_full is None or detail_full.empty:
        return detail_full
    df = detail_full.copy()
    for item, g in df.groupby("Item", sort=False):
        zero = g[(pd.to_numeric(g["Average of SS Wk3"], errors="coerce").fillna(0) <= 0)
                 & (g["Hard Lock"] == False)
                 & (g["Priority Rule Mode"].fillna("").astype(str) == "")]
        if len(zero) < 2:
            continue
        idxs = list(zero.index)
        guard = 0
        while guard < 100000:
            guard += 1
            si = df.loc[idxs, "Current SI After"].astype(float)
            donor = si.idxmax()
            recipient = si.idxmin()
            gap = float(si.loc[donor] - si.loc[recipient])
            if gap <= 1:
                break
            donor_floor = int(max(0, pd.to_numeric(pd.Series([df.at[donor, "Future Firm Floor"] if "Future Firm Floor" in df.columns else 0]), errors="coerce").fillna(0).iloc[0]))
            if int(df.at[donor, "F Wk3 After Destination Change"]) <= donor_floor:
                # This donor cannot give another unit without violating its future-demand protection floor.
                break
            df.at[donor, "F Wk3 After Destination Change"] -= 1
            df.at[recipient, "F Wk3 After Destination Change"] += 1
            df.at[donor, "Net Destination Change"] -= 1
            df.at[recipient, "Net Destination Change"] += 1
            df.at[donor, "Current SI After"] -= 1
            df.at[recipient, "Current SI After"] += 1
            df.at[donor, "SS % After"] = safe_ss_ratio(float(df.at[donor, "Current SI After"]), 0.0)
            df.at[recipient, "SS % After"] = safe_ss_ratio(float(df.at[recipient, "Current SI After"]), 0.0)
    # Validate conservation item-by-item.
    before = df.groupby("Item")["F Wk3 Original"].sum()
    after = df.groupby("Item")["F Wk3 After Destination Change"].sum()
    if not before.equals(after):
        raise ValueError("Zero-SS equalization changed total Firm PO.")
    return df


def _allocate_secondary_vendor_greedy(detail_full: pd.DataFrame, priority_rules: Optional[Dict[str, PriorityRule]] = None, respect_priority_rank: bool = False, freeze_hold_buy: bool = False) -> pd.DataFrame:
    """Mirror main-vendor allocation logic for sub-vendor supply.

    Only the input quantity changes: Other Vendor Supply becomes the Firm PO pool, while
    the baseline Current SI is the Main Vendor SI After result.
    """
    df = detail_full.copy()
    if df.empty:
        return df
    if "Other Vendor Supply" not in df.columns:
        df["Other Vendor Supply"] = 0.0
    df["Other Vendor Supply"] = pd.to_numeric(df["Other Vendor Supply"], errors="coerce").fillna(0.0)
    priority_rules = priority_rules or {}
    all_records = []
    for _, g0 in df.groupby("Item", sort=True):
        g = g0.copy().reset_index()
        orig = np.rint(g["Other Vendor Supply"].to_numpy(dtype=float)).astype(int)
        temp = g.copy()
        temp["F Wk3"] = orig
        temp["Current SI"] = pd.to_numeric(temp["Current SI After"], errors="coerce").fillna(0.0)
        temp["Average of SS Wk3"] = pd.to_numeric(temp["Average of SS Wk3"], errors="coerce").fillna(0.0)
        temp["Firm PO Total"] = int(orig.sum())
        if "Sub Future Firm Floor" in g.columns:
            temp["Future Firm Floor"] = pd.to_numeric(g["Sub Future Firm Floor"], errors="coerce").fillna(0.0)
        # Skip redistribution if all sub-vendor supply is zero: keep zeros and preserve main result.
        if int(orig.sum()) == 0:
            final = orig.copy()
        else:
            item_whse = set(temp["Whse"].astype(str).tolist())
            item_rules = {wh: rule for wh, rule in priority_rules.items() if wh in item_whse}
            allocated = allocate_item(temp, item_rules, respect_priority_rank=respect_priority_rank, freeze_hold_buy=freeze_hold_buy)
            final = pd.to_numeric(allocated["F Wk3 After Destination Change"], errors="coerce").fillna(0).astype(int).to_numpy()
        for i, row_index in enumerate(g["index"]):
            all_records.append((int(row_index), int(orig[i]), int(final[i])))
    sub = pd.DataFrame(all_records, columns=["_idx", "Sub Vendor F Original", "Sub Vendor F After Destination Change"]).set_index("_idx").sort_index()
    df = df.join(sub, how="left")
    for c in ["Sub Vendor F Original", "Sub Vendor F After Destination Change"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
    df["Sub Vendor Net Destination Change"] = df["Sub Vendor F After Destination Change"] - df["Sub Vendor F Original"]
    # The sub-vendor flow starts from the completed main-vendor result.
    df["Sub Vendor SI Before"] = pd.to_numeric(df["Current SI After"], errors="coerce").fillna(0.0)
    df["Sub Vendor SI After"] = df["Sub Vendor SI Before"] + df["Sub Vendor Net Destination Change"]
    df["Sub Vendor SS% After"] = df.apply(lambda r: safe_ss_ratio(float(r["Sub Vendor SI After"]), float(r["Average of SS Wk3"])), axis=1)
    df["Sub Vendor DC Note"] = df["Sub Vendor F Original"].map(lambda x: "Suggestion only" if x > 0 else "")
    return df


def _sum_preserving_round(values: np.ndarray, total: int, lower: Optional[np.ndarray] = None, upper: Optional[np.ndarray] = None) -> np.ndarray:
    values = np.asarray(values, dtype=float)
    n = len(values)
    lower = np.zeros(n, dtype=int) if lower is None else np.asarray(lower, dtype=int)
    upper = np.full(n, max(total, 0), dtype=int) if upper is None else np.asarray(upper, dtype=int)
    x = np.floor(values).astype(int)
    x = np.clip(x, lower, upper)
    diff = int(total - int(x.sum()))
    frac = values - np.floor(values)
    guard = 0
    while diff != 0 and guard < 100000:
        guard += 1
        if diff > 0:
            candidates = [i for i in range(n) if x[i] < upper[i]]
            if not candidates:
                break
            candidates.sort(key=lambda i: frac[i], reverse=True)
            for i in candidates:
                if diff == 0: break
                x[i] += 1; diff -= 1
        else:
            candidates = [i for i in range(n) if x[i] > lower[i]]
            if not candidates:
                break
            candidates.sort(key=lambda i: frac[i])
            for i in candidates:
                if diff == 0: break
                x[i] -= 1; diff += 1
    return x



def prepare_optimizer_input(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    required = ["Item", "ProdResourceID", "Whse", "F Wk3", "Sum of SI Wk3", "Average of SS Wk3"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Optimizer input is missing required columns: {missing}")
    if not [c for c in df.columns if "vendor" in c.lower()]:
        df[VENDOR_FALLBACK_COL] = ""

    df["Item"] = df["Item"].map(normalize_item)
    df = df[df["Item"] != ""].copy()
    df["Whse"] = df["Whse"].map(normalize_whse)
    for c in ["F Wk3", "Sum of SI Wk3", "Average of SS Wk3"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    if "Sum of SI-SS Wk3" in df.columns:
        df["Sum of SI-SS Wk3"] = pd.to_numeric(df["Sum of SI-SS Wk3"], errors="coerce").fillna(0)
    else:
        df["Sum of SI-SS Wk3"] = pd.NA

    # Business rule update:
    # Main vendor F Wk3 is eligible for optimizer allocation. Confirmed other vendor supply
    # and Firm PO Reconciliation Gap update SI/SS only.
    if "Main Vendor F Wk3" not in df.columns:
        df["Main Vendor F Wk3"] = df["F Wk3"]
    if "Other Vendor Supply" not in df.columns:
        df["Other Vendor Supply"] = 0.0
    if "Firm PO Reconciliation Gap" not in df.columns:
        df["Firm PO Reconciliation Gap"] = 0.0
    if "PSW F Used for Reconciliation" not in df.columns:
        df["PSW F Used for Reconciliation"] = pd.to_numeric(df["Main Vendor F Wk3"], errors="coerce").fillna(0.0) + pd.to_numeric(df["Other Vendor Supply"], errors="coerce").fillna(0.0)
    for _c in ["Main Vendor F Wk3", "Other Vendor Supply", "Firm PO Reconciliation Gap", "PSW F Used for Reconciliation"]:
        df[_c] = pd.to_numeric(df[_c], errors="coerce").fillna(0.0)

    supply_already_applied = df.get("_supply_already_applied", pd.Series(False, index=df.index)).fillna(False).astype(bool)
    if not supply_already_applied.all():
        df.loc[~supply_already_applied, "Total Supply Added to SI"] = (
            df.loc[~supply_already_applied, "Main Vendor F Wk3"]
            + df.loc[~supply_already_applied, "Other Vendor Supply"]
            + df.loc[~supply_already_applied, "Firm PO Reconciliation Gap"]
        )
        df.loc[~supply_already_applied, "Original SI Before Supply"] = df.loc[~supply_already_applied, "Sum of SI Wk3"]
        df.loc[~supply_already_applied, "Original SI-SS Before Supply"] = df.loc[~supply_already_applied, "Sum of SI-SS Wk3"]
        df.loc[~supply_already_applied, "New SI"] = df.loc[~supply_already_applied, "Original SI Before Supply"] + df.loc[~supply_already_applied, "Total Supply Added to SI"]
        df.loc[~supply_already_applied, "New SI-SS"] = df.loc[~supply_already_applied, "Original SI-SS Before Supply"] + df.loc[~supply_already_applied, "Total Supply Added to SI"]
        df.loc[~supply_already_applied, "Sum of SI Wk3"] = df.loc[~supply_already_applied, "New SI"]
        df.loc[~supply_already_applied, "Sum of SI-SS Wk3"] = df.loc[~supply_already_applied, "New SI-SS"]

    df["Total Supply Added to SI"] = pd.to_numeric(df.get("Total Supply Added to SI", 0), errors="coerce").fillna(0.0)
    df["Original SI Before Supply"] = pd.to_numeric(df.get("Original SI Before Supply", df["Sum of SI Wk3"]), errors="coerce").fillna(0.0)
    df["Original SI-SS Before Supply"] = pd.to_numeric(df.get("Original SI-SS Before Supply", df["Sum of SI-SS Wk3"]), errors="coerce").fillna(0.0)
    df["New SI"] = pd.to_numeric(df.get("New SI", df["Sum of SI Wk3"]), errors="coerce").fillna(0.0)
    df["New SI-SS"] = pd.to_numeric(df.get("New SI-SS", df["Sum of SI-SS Wk3"]), errors="coerce").fillna(0.0)
    df["Sum of SI Wk3"] = pd.to_numeric(df["Sum of SI Wk3"], errors="coerce").fillna(0.0)
    df["Sum of SI-SS Wk3"] = pd.to_numeric(df["Sum of SI-SS Wk3"], errors="coerce").fillna(0.0)

    df["Current SI"] = df["Sum of SI Wk3"]
    df["Current SS%"] = df.apply(lambda r: safe_ss_ratio(float(r["Current SI"]), float(r["Average of SS Wk3"])), axis=1)
    # Always recompute Firm PO Total from the current optimizer input.
    # This avoids merge collisions such as Firm PO Total_x / Firm PO Total_y
    # when the optional auto-balancing path has already created the column.
    df["Firm PO Total"] = df.groupby("Item")["F Wk3"].transform("sum")
    df["Firm PO Total"] = pd.to_numeric(df["Firm PO Total"], errors="coerce").fillna(0.0)
    return df


def attach_output_week_metrics(
    detail_full: pd.DataFrame,
    firm_week_snapshot: pd.DataFrame,
    firm_week: date,
) -> pd.DataFrame:
    """Attach clear Firm-PO-Week and Balance-Week SI/SS audit metrics for final output.

    Internal optimizer column names are intentionally preserved so the allocation logic and
    downstream UPLOAD/OSQP behavior do not change. This helper only adds user-facing audit aliases.

    User-facing definitions:
      Current SI / Current SI-SS / Current SS / Current SS%
        = Firm PO Week, before any destination-change redistribution.
      SI After DC @ Firm PO Week
        = Firm-week Current SI plus Main + Sub vendor destination-change deltas.
      Balance-week metrics
        = selected Balance Week, rebuilt fresh from PlanDetailTimeline; After DC includes
          both Main and Sub vendor destination-change deltas.
    """
    if detail_full is None or detail_full.empty:
        return detail_full

    df = detail_full.copy()
    idx = df.index

    # Preserve the selected Balance Week metrics before adding Firm-week aliases.
    bal_si = pd.to_numeric(df.get("Current SI", pd.Series(0.0, index=idx)), errors="coerce").fillna(0.0)
    bal_ss = pd.to_numeric(df.get("Average of SS Wk3", pd.Series(0.0, index=idx)), errors="coerce").fillna(0.0)
    main_net = pd.to_numeric(df.get("Net Destination Change", pd.Series(0.0, index=idx)), errors="coerce").fillna(0.0)
    sub_net = pd.to_numeric(df.get("Sub Vendor Net Destination Change", pd.Series(0.0, index=idx)), errors="coerce").fillna(0.0)
    total_net = main_net + sub_net

    df["SI Before DC @ Balance Week"] = bal_si
    df["SI-SS Before DC @ Balance Week"] = bal_si - bal_ss
    df["Effective SS @ Balance Week"] = bal_ss
    df["SS% Before DC @ Balance Week"] = [safe_ss_ratio(float(si), float(ss)) for si, ss in zip(bal_si, bal_ss)]
    df["SI After DC @ Balance Week"] = bal_si + total_net
    df["SI-SS After DC @ Balance Week"] = df["SI After DC @ Balance Week"] - bal_ss
    df["SS% After DC @ Balance Week"] = [
        safe_ss_ratio(float(si), float(ss))
        for si, ss in zip(df["SI After DC @ Balance Week"], bal_ss)
    ]

    # Build Firm PO Week baseline from a fresh firm-week candidate snapshot.
    snap = firm_week_snapshot.copy() if firm_week_snapshot is not None else pd.DataFrame()
    if not snap.empty:
        keep = [c for c in ["Item", "Whse", "Current SI", "Average of SS Wk3"] if c in snap.columns]
        snap = snap[keep].copy()
        snap["Item"] = snap["Item"].map(normalize_item)
        snap["Whse"] = snap["Whse"].map(normalize_whse)
        snap = snap.drop_duplicates(["Item", "Whse"])
        snap = snap.rename(columns={
            "Current SI": "_Firm Week Current SI",
            "Average of SS Wk3": "_Firm Week Current SS",
        })
        df = df.merge(snap, on=["Item", "Whse"], how="left", sort=False)
    else:
        df["_Firm Week Current SI"] = pd.NA
        df["_Firm Week Current SS"] = pd.NA

    # If a rare source row cannot be rebuilt at Firm Week, use Balance metrics only as a safe audit fallback.
    # The fallback is visible through the row's Firm/Balance Week columns and does not alter allocation.
    firm_si = pd.to_numeric(df.get("_Firm Week Current SI"), errors="coerce")
    firm_ss = pd.to_numeric(df.get("_Firm Week Current SS"), errors="coerce")
    firm_si = firm_si.where(firm_si.notna(), pd.to_numeric(df["SI Before DC @ Balance Week"], errors="coerce"))
    firm_ss = firm_ss.where(firm_ss.notna(), pd.to_numeric(df["Effective SS @ Balance Week"], errors="coerce"))

    df["_Firm Week Current SI"] = firm_si
    df["_Firm Week Current SS"] = firm_ss
    df["_Firm Week Current SI-SS"] = firm_si - firm_ss
    df["_Firm Week Current SS%"] = [safe_ss_ratio(float(si), float(ss)) for si, ss in zip(firm_si, firm_ss)]
    df["_Firm Week SI After DC"] = firm_si + total_net
    df["_Firm Week SI-SS After DC"] = df["_Firm Week SI After DC"] - firm_ss
    df["_Firm Week SS% After DC"] = [
        safe_ss_ratio(float(si), float(ss))
        for si, ss in zip(df["_Firm Week SI After DC"], firm_ss)
    ]
    df["Firm PO Week"] = firm_week
    return df


def build_detail_output(detail: pd.DataFrame) -> pd.DataFrame:
    if detail is None or detail.empty:
        return pd.DataFrame()
    detail = detail.copy()

    # Direct run_optimizer() compatibility: when the explicit output audit layer has not yet
    # been attached, treat the current optimizer week as both Firm PO Week and Balance Week.
    # process_files() later replaces these aliases with the true Firm-week snapshot when Auto Separate is used.
    idx = detail.index
    _curr_si = pd.to_numeric(detail.get("Current SI", pd.Series(0.0, index=idx)), errors="coerce").fillna(0.0)
    _curr_ss = pd.to_numeric(detail.get("Average of SS Wk3", pd.Series(0.0, index=idx)), errors="coerce").fillna(0.0)
    _main_net = pd.to_numeric(detail.get("Net Destination Change", pd.Series(0.0, index=idx)), errors="coerce").fillna(0.0)
    _sub_net = pd.to_numeric(detail.get("Sub Vendor Net Destination Change", pd.Series(0.0, index=idx)), errors="coerce").fillna(0.0)
    _total_net = _main_net + _sub_net
    if "_Firm Week Current SI" not in detail.columns:
        detail["_Firm Week Current SI"] = _curr_si
        detail["_Firm Week Current SS"] = _curr_ss
        detail["_Firm Week Current SI-SS"] = _curr_si - _curr_ss
        detail["_Firm Week Current SS%"] = [safe_ss_ratio(float(si), float(ss)) for si, ss in zip(_curr_si, _curr_ss)]
        detail["_Firm Week SI After DC"] = _curr_si + _total_net
        detail["_Firm Week SI-SS After DC"] = detail["_Firm Week SI After DC"] - _curr_ss
        detail["_Firm Week SS% After DC"] = [safe_ss_ratio(float(si), float(ss)) for si, ss in zip(detail["_Firm Week SI After DC"], _curr_ss)]
    if "SI Before DC @ Balance Week" not in detail.columns:
        detail["SI Before DC @ Balance Week"] = _curr_si
        detail["SI-SS Before DC @ Balance Week"] = _curr_si - _curr_ss
        detail["Effective SS @ Balance Week"] = _curr_ss
        detail["SS% Before DC @ Balance Week"] = [safe_ss_ratio(float(si), float(ss)) for si, ss in zip(_curr_si, _curr_ss)]
        detail["SI After DC @ Balance Week"] = _curr_si + _total_net
        detail["SI-SS After DC @ Balance Week"] = detail["SI After DC @ Balance Week"] - _curr_ss
        detail["SS% After DC @ Balance Week"] = [safe_ss_ratio(float(si), float(ss)) for si, ss in zip(detail["SI After DC @ Balance Week"], _curr_ss)]

    # Keep the familiar v6.x structure as much as possible. The first SI/SS block now refers
    # explicitly to Firm PO Week ("Current" for end users), followed by the selected Balance Week block.
    preferred = [
        "Item", "ProdResourceID", "Whse", "F Wk3",
        "_Firm Week Current SI", "_Firm Week Current SI-SS", "_Firm Week Current SS", "_Firm Week Current SS%",
        "Firm PO Total", "F Wk3 Original", "F Wk3 After Destination Change", "Net Destination Change",
        "_Firm Week SI After DC", "_Firm Week SI-SS After DC", "_Firm Week SS% After DC",
        "SI Before DC @ Balance Week", "SI-SS Before DC @ Balance Week", "Effective SS @ Balance Week",
        "SS% Before DC @ Balance Week", "SI After DC @ Balance Week", "SI-SS After DC @ Balance Week",
        "SS% After DC @ Balance Week",
        "Remaining Unallocated PO",
        "Priority Rule Mode", "Priority Rule Value", "Priority Rank", "Priority Target F After", "Hard Lock",
        "Original SI Before Supply", "Original SI-SS Before Supply", "New SI", "New SI-SS",
        "Main Vendor F Wk3", "Other Vendor Supply", "Timeline Firm PO", "PSW F Used for Reconciliation",
        "Firm PO Reconciliation Gap", "Total Supply Added to SI", "Other Vendor List",
        "Sub Vendor F Original", "Sub Vendor F After Destination Change", "Sub Vendor Net Destination Change",
        "Sub Vendor SI Before", "Sub Vendor SI After", "Sub Vendor SS% After", "Sub Vendor DC Note",
        "Main Vendor", "Vendor", "Firm PO Week", "Balance Week", "Hold/ Buy",
    ]
    final_cols = [c for c in preferred if c in detail.columns]
    final_cols += [c for c in PLAN_METADATA_COLUMNS if c in detail.columns and c not in final_cols]
    out = detail[final_cols].copy()
    rename_map = {
        "F Wk3": "Firm PO",
        "_Firm Week Current SI": "Current SI",
        "_Firm Week Current SI-SS": "Current SI-SS",
        "_Firm Week Current SS": "Current SS",
        "_Firm Week Current SS%": "Current SS%",
        "F Wk3 Original": "Firm PO Original",
        "F Wk3 After Destination Change": "Firm PO After Destination Change",
        "_Firm Week SI After DC": "SI After DC @ Firm PO Week",
        "_Firm Week SI-SS After DC": "SI-SS After DC @ Firm PO Week",
        "_Firm Week SS% After DC": "SS% After DC @ Firm PO Week",
    }
    out = out.rename(columns=rename_map)
    return out

def build_summary(detail: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, g in detail.groupby("Item", sort=True):
        rows.append({
            "Item": g["Item"].iloc[0],
            "ProdResourceID": g["ProdResourceID"].iloc[0],
            "Firm PO Total": int(g["Firm PO Total"].iloc[0]),
            "Total F Before": int(g["F Wk3 Original"].sum()),
            "Total F After": int(g["F Wk3 After Destination Change"].sum()),
            "Min SI After": int(g["Current SI After"].min()),
            "Max SI After": int(g["Current SI After"].max()),
            "Min SS % After": float(g["SS % After"].min()),
            "Max SS % After": float(g["SS % After"].max()),
        })
    return pd.DataFrame(rows)


def _sum_preserving_round(values: np.ndarray, total: int, lower: Optional[np.ndarray] = None, upper: Optional[np.ndarray] = None) -> np.ndarray:
    values = np.asarray(values, dtype=float)
    n = len(values)
    if lower is None:
        lower = np.zeros(n, dtype=int)
    else:
        lower = np.asarray(lower, dtype=int)
    if upper is None:
        upper = np.full(n, max(total, 0), dtype=int)
    else:
        upper = np.asarray(upper, dtype=int)
    x = np.floor(values).astype(int)
    x = np.clip(x, lower, upper)
    diff = int(round(total - int(x.sum())))
    frac = values - np.floor(values)
    guard = 0
    while diff != 0 and guard < 100000:
        guard += 1
        if diff > 0:
            candidates = [i for i in range(n) if x[i] < upper[i]]
            if not candidates:
                break
            candidates.sort(key=lambda i: frac[i], reverse=True)
            for i in candidates:
                if diff == 0:
                    break
                if x[i] < upper[i]:
                    x[i] += 1
                    diff -= 1
        else:
            candidates = [i for i in range(n) if x[i] > lower[i]]
            if not candidates:
                break
            candidates.sort(key=lambda i: frac[i])
            for i in candidates:
                if diff == 0:
                    break
                if x[i] > lower[i]:
                    x[i] -= 1
                    diff += 1
    return x


def _osqp_equalize_single_item(F_orig: np.ndarray, curr_si: np.ndarray, avg_ss: np.ndarray) -> Tuple[Optional[np.ndarray], str]:
    """OSQP second-pass optimizer: minimize SS% spread around the feasible network coverage.

    It preserves total F and keeps x >= 0. If OSQP/scipy is unavailable or the item is not applicable,
    returns (None, reason).
    """
    try:
        import osqp  # type: ignore
        from scipy import sparse  # type: ignore
    except Exception as exc:
        return None, f"OSQP unavailable: {exc}"

    # Work with integer shipment quantities. This is important because the output
    # Net Destination Change is calculated against integer Original F.
    # Using round(sum(F_orig)) can create a non-zero total net when individual
    # rows contain decimals. Use sum(round(each row)) instead.
    F_orig = np.asarray(F_orig, dtype=float)
    F_orig = np.rint(F_orig).astype(float)
    curr_si = np.asarray(curr_si, dtype=float)
    avg_ss = np.asarray(avg_ss, dtype=float)
    n = len(F_orig)
    total_f = int(np.sum(F_orig))
    if n <= 1 or total_f <= 0:
        return None, "Not applicable: one warehouse or zero total F"

    valid = avg_ss > 0
    if not np.any(valid):
        return None, "Not applicable: all SS are zero"

    # Option B: equalize coverage between participating warehouses, not necessarily to 100%.
    # Feasible target coverage is based on the network after preserving total F.
    total_final_si = float(np.sum(curr_si + (F_orig * 0)))  # sum of current SI before reallocation
    target_ratio = total_final_si / float(np.sum(avg_ss[valid])) if float(np.sum(avg_ss[valid])) != 0 else 0.0

    k = np.zeros(n)
    c = np.zeros(n)
    weights = np.ones(n)
    mean_ss = np.mean(avg_ss[valid]) if np.any(valid) else 1.0
    for i in range(n):
        if avg_ss[i] > 0:
            k[i] = 1.0 / avg_ss[i]
            c[i] = (curr_si[i] - F_orig[i]) / avg_ss[i] - target_ratio
            weights[i] = min(max(avg_ss[i] / mean_ss, 0.1), 10.0)
        else:
            # Avoid unstable zero-SS rows. Movement penalty still keeps them reasonable.
            k[i] = 0.0
            c[i] = 0.0
            weights[i] = 0.1

    alpha_move = 1e-5
    P_diag = 2.0 * weights * (k ** 2) + 2.0 * alpha_move
    q = 2.0 * weights * k * c - 2.0 * alpha_move * F_orig
    P = sparse.diags(P_diag, format="csc")
    A = sparse.vstack([
        sparse.csr_matrix(np.ones((1, n))),
        sparse.eye(n, format="csc"),
    ], format="csc")
    l = np.concatenate([[total_f], np.zeros(n)])
    u = np.concatenate([[total_f], np.full(n, max(total_f, int(np.max(F_orig) * 2) + total_f + 1))])

    prob = osqp.OSQP()
    prob.setup(P=P, q=q, A=A, l=l, u=u, verbose=False, eps_abs=1e-5, eps_rel=1e-5, max_iter=30000, polish=True)
    res = prob.solve()
    if res.info.status_val not in (1, 2) or res.x is None:
        return None, f"OSQP failed: {res.info.status}"
    x_int = _sum_preserving_round(res.x[:n], total_f, lower=np.zeros(n, dtype=int))
    return x_int, "OSQP Equalize SS%"


def build_osqp_sheets(detail_full: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """Build optional OSQP sheets for main vendor and sub/other vendor.

    Main vendor OSQP starts from New SI before current/main destination change.
    Sub vendor OSQP starts after the OSQP main-vendor result when available; otherwise after current main result.
    """
    sheets: Dict[str, pd.DataFrame] = {}
    if detail_full is None or detail_full.empty:
        return sheets

    main_records = []
    sub_records = []
    main_si_after_by_index = {}

    for _, g0 in detail_full.groupby("Item", sort=True):
        g = g0.copy().reset_index()
        f_orig = pd.to_numeric(g["F Wk3 Original"], errors="coerce").fillna(0).to_numpy(dtype=float)
        curr_si = pd.to_numeric(g["Current SI"], errors="coerce").fillna(0).to_numpy(dtype=float)
        avg_ss = pd.to_numeric(g["Average of SS Wk3"], errors="coerce").fillna(0).to_numpy(dtype=float)
        x, status = _osqp_equalize_single_item(f_orig, curr_si, avg_ss)
        if x is None:
            x = pd.to_numeric(g["F Wk3 After Destination Change"], errors="coerce").fillna(0).to_numpy(dtype=int)
        for i in range(len(g)):
            net = int(x[i]) - int(round(f_orig[i]))
            si_after = float(curr_si[i]) + net
            main_si_after_by_index[int(g.loc[i, "index"])] = si_after
            main_records.append({
                "Item": g.loc[i, "Item"],
                "ProdResourceID": g.loc[i, "ProdResourceID"],
                "Whse": g.loc[i, "Whse"],
                "Vendor": g.loc[i, "Vendor"] if "Vendor" in g.columns else "",
                "F Wk3 Original": int(round(f_orig[i])),
                "OSQP F Wk3 After Destination Change": int(x[i]),
                "OSQP Net Destination Change": net,
                "OSQP SI After": si_after,
                "OSQP SS% After": safe_ss_ratio(si_after, float(avg_ss[i])),
                "Average of SS Wk3": float(avg_ss[i]),
                "OSQP Method/Status": status,
            })

        sub_orig = pd.to_numeric(g["Other Vendor Supply"], errors="coerce").fillna(0).to_numpy(dtype=float) if "Other Vendor Supply" in g.columns else np.zeros(len(g))
        # Use integer sub-vendor original quantities for both OSQP constraint and net-change reporting.
        # This guarantees SUM(OSQP Sub Vendor Net Destination Change) = 0 by item.
        sub_orig_int = np.rint(sub_orig).astype(int)
        sub_curr_si = np.array([main_si_after_by_index.get(int(g.loc[i, "index"]), float(g.loc[i, "Current SI After"])) for i in range(len(g))], dtype=float)
        sx, sstatus = _osqp_equalize_single_item(sub_orig_int, sub_curr_si, avg_ss)
        if sx is None:
            sx = pd.to_numeric(g.get("Sub Vendor F After Destination Change", pd.Series(np.zeros(len(g)))), errors="coerce").fillna(0).to_numpy(dtype=int)
        sx = np.asarray(sx, dtype=int)
        # Last safety check: force total sub-vendor after quantity to equal total original quantity.
        sx = _sum_preserving_round(sx.astype(float), int(sub_orig_int.sum()), lower=np.zeros(len(sx), dtype=int))
        for i in range(len(g)):
            snet = int(sx[i]) - int(sub_orig_int[i])
            ssi_after = float(sub_curr_si[i]) + snet
            sub_records.append({
                "Item": g.loc[i, "Item"],
                "ProdResourceID": g.loc[i, "ProdResourceID"],
                "Whse": g.loc[i, "Whse"],
                "Other Vendor List": g.loc[i, "Other Vendor List"] if "Other Vendor List" in g.columns else "",
                "Sub Vendor F Original": int(sub_orig_int[i]),
                "OSQP Sub Vendor SI Before": float(sub_curr_si[i]),
                "OSQP Sub Vendor F After Destination Change": int(sx[i]),
                "OSQP Sub Vendor Net Destination Change": snet,
                "OSQP Sub Vendor SI After": ssi_after,
                "OSQP Sub Vendor SS% After": safe_ss_ratio(ssi_after, float(avg_ss[i])),
                "Average of SS Wk3": float(avg_ss[i]),
                "OSQP Method/Status": sstatus,
            })

    sheets["OSQP Main Vendor"] = pd.DataFrame(main_records)
    sheets["OSQP Sub Vendor"] = pd.DataFrame(sub_records)
    return sheets


def run_optimizer(
    optimizer_input: pd.DataFrame,
    priority_rules: Optional[Dict[str, PriorityRule]] = None,
    respect_priority_rank: bool = False,
    freeze_hold_buy: bool = False,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    priority_rules = priority_rules or {}
    df = prepare_optimizer_input(optimizer_input)
    records = []
    for _, group in df.groupby("Item", sort=True):
        item_whse = set(group["Whse"].astype(str).tolist())
        item_rules = {wh: rule for wh, rule in priority_rules.items() if wh in item_whse}
        allocated = allocate_item(group.copy(), item_rules, respect_priority_rank=respect_priority_rank, freeze_hold_buy=freeze_hold_buy)
        records.extend(allocated.to_dict("records"))
    detail_full = pd.DataFrame.from_records(records) if records else pd.DataFrame()
    if detail_full.empty:
        return pd.DataFrame(), pd.DataFrame(), detail_full
    # Fallback equalization for zero-SS groups is always automatic and is applied after the main allocation.
    detail_full = apply_zero_ss_equalization(detail_full)

    # Final validation after post-processing: FIRM=0 rows must remain zero.
    if "Firm Zero Target" in detail_full.columns:
        bad_firm_zero = detail_full[
            (detail_full["Firm Zero Target"] == True)
            & (detail_full["Hard Lock"] == False)
            & (pd.to_numeric(detail_full["F Wk3 After Destination Change"], errors="coerce").fillna(0) != 0)
        ]
        if not bad_firm_zero.empty:
            raise ValueError("FIRM=0 priority validation failed: targeted warehouse still has Firm PO after allocation.")

    # Mirror the same priority logic for sub-vendor suggestions after main-vendor optimization.
    detail_full = _allocate_secondary_vendor_greedy(detail_full, priority_rules=priority_rules, respect_priority_rank=respect_priority_rank, freeze_hold_buy=freeze_hold_buy)
    if int(detail_full["F Wk3 Original"].sum()) != int(detail_full["F Wk3 After Destination Change"].sum()):
        raise ValueError("Total Firm PO is not preserved after optimization.")
    detail = build_detail_output(detail_full)
    summary = build_summary(detail_full)
    return detail, summary, detail_full


# ============================================================
# Output writer
# ============================================================

def autofit(ws):
    for col_cells in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            value = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(value))
        ws.column_dimensions[col_letter].width = min(max_len + 2, 45)


def style_sheet(ws):
    ws.freeze_panes = "A2"
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    autofit(ws)



def _split_vendor_list(value: str) -> List[str]:
    if value is None:
        return []
    text = str(value).strip()
    if not text:
        return []
    return [normalize_vendor(x) for x in re.split(r"[,;|]+", text) if str(x).strip()]


def build_upload_sheet(detail_full: pd.DataFrame, firm_week: date, current_week: date, psw_vendor_detail: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    cols = ["Vendor", "ProductionResource", "Item", "WarehouseCode", "WeekIndex", "Quantity"]
    if detail_full is None or detail_full.empty:
        return pd.DataFrame(columns=cols)
    week_index = max(1, int(round((firm_week - current_week).days / 7.0)) + 1)

    # Original sub-vendor quantities by Item + Whse + Vendor at Firm PO Week, for proportional split only when needed.
    vendor_share = {}
    if psw_vendor_detail is not None and not psw_vendor_detail.empty:
        x = psw_vendor_detail.copy()
        if "Vendor Role" in x.columns:
            x = x[x["Vendor Role"].astype(str).str.upper().eq("OTHER")]
        if "PSW Week" in x.columns:
            wk = pd.to_datetime(x["PSW Week"], errors="coerce").dt.date
            x = x[wk.eq(firm_week)]
        if not x.empty and "PSW Quantity" in x.columns:
            x["Item"] = x["Item"].map(normalize_item)
            x["Whse"] = x["Whse"].map(normalize_whse)
            x["Vendor"] = x["Vendor"].map(normalize_vendor)
            x["PSW Quantity"] = pd.to_numeric(x["PSW Quantity"], errors="coerce").fillna(0.0)
            g = x.groupby(["Item","Whse","Vendor"], dropna=False)["PSW Quantity"].sum().reset_index()
            for (item,whse), gg in g.groupby(["Item","Whse"], dropna=False):
                total=float(gg["PSW Quantity"].sum())
                if total>0:
                    vendor_share[(str(item),str(whse))]=[(str(v),float(q)/total) for v,q in zip(gg["Vendor"],gg["PSW Quantity"]) if str(v).strip()]

    rows=[]
    # Main vendor rows first. Only non-zero main destination changes.
    main_net=pd.to_numeric(detail_full["Net Destination Change"], errors="coerce").fillna(0)
    for _,r in detail_full.loc[main_net.ne(0)].iterrows():
        vendor=normalize_vendor(r.get("Main Vendor") or r.get("Vendor") or "")
        qty=int(round(float(r.get("F Wk3 After Destination Change",0) or 0)))
        if not vendor:
            continue
        rows.append({"Vendor":vendor,"ProductionResource":str(r.get("ProdResourceID") or "").strip(),"Item":normalize_item(r.get("Item")),"WarehouseCode":normalize_whse(r.get("Whse")),"WeekIndex":week_index,"Quantity":qty,"_role":0})

    # Sub vendor rows after main vendor rows. Only non-zero sub destination changes.
    if "Sub Vendor Net Destination Change" in detail_full.columns:
        sub_net=pd.to_numeric(detail_full["Sub Vendor Net Destination Change"], errors="coerce").fillna(0)
    else:
        sub_net=pd.Series(0, index=detail_full.index, dtype="float64")
    for _,r in detail_full.loc[sub_net.ne(0)].iterrows():
        item=normalize_item(r.get("Item")); whse=normalize_whse(r.get("Whse"))
        qty=int(round(float(r.get("Sub Vendor F After Destination Change",0) or 0)))
        vendors=_split_vendor_list(r.get("Other Vendor List", ""))
        if not vendors:
            continue
        shares=vendor_share.get((str(item),str(whse)))
        if shares:
            rawq=[qty*share for _,share in shares]
            ints=[int(q) for q in rawq]
            rem=qty-sum(ints)
            order=sorted(range(len(rawq)), key=lambda i: rawq[i]-ints[i], reverse=True)
            for i in order[:max(0,rem)]: ints[i]+=1
            # Ensure every listed vendor is present in shares; vendors without original qty get zero.
            share_map=dict(shares)
            final_qty=[ints[i] for i in range(len(ints))]
            ordered_vendors=[v for v,_ in shares]
            for vendor,q in zip(ordered_vendors, final_qty):
                # Keep zero-quantity donor rows when Net Destination Change is non-zero.
                rows.append({"Vendor":vendor,"ProductionResource":str(r.get("ProdResourceID") or "").strip(),"Item":item,"WarehouseCode":whse,"WeekIndex":week_index,"Quantity":int(q),"_role":1})
        else:
            rows.append({"Vendor":vendors[0],"ProductionResource":str(r.get("ProdResourceID") or "").strip(),"Item":item,"WarehouseCode":whse,"WeekIndex":week_index,"Quantity":qty,"_role":1})

    if not rows:
        return pd.DataFrame(columns=cols)
    out=pd.DataFrame(rows)
    # Do NOT filter Quantity=0. A donor with non-zero destination change can legitimately end at zero;
    # the zero upload row is required to clear the old Firm PO on the target system.
    out=out.sort_values(["_role","Item","WarehouseCode","Vendor"], kind="stable")[cols].reset_index(drop=True)
    return out

def write_excel_output(
    output_path: str,
    detail: pd.DataFrame,
    summary: pd.DataFrame,
    optimizer_input: pd.DataFrame,
    debug_sheets: Optional[Dict[str, pd.DataFrame]] = None,
    extra_sheets: Optional[Dict[str, pd.DataFrame]] = None,
    upload_sheet: Optional[pd.DataFrame] = None,
) -> str:
    output_path = ensure_unique_output_path(output_path)

    # No output-level INF/-INF sanitization: balance logic handles zero-SS planning
    # via Last Positive SS bridge and Runout Balance Mode instead of masking results.
    detail_out = detail.copy() if detail is not None else pd.DataFrame()
    upload_out = upload_sheet.copy() if upload_sheet is not None else None
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        detail_out.to_excel(writer, sheet_name="Optimized Data", index=False)
        if upload_out is not None:
            upload_out.to_excel(writer, sheet_name="UPLOAD", index=False)
        for name, df in (extra_sheets or {}).items():
            if df is not None and not df.empty:
                df.to_excel(writer, sheet_name=name[:31], index=False)
        wb = writer.book
        for ws in wb.worksheets:
            style_sheet(ws)
            if ws.title == "UPLOAD":
                ws.auto_filter.ref = f"A1:F{max(ws.max_row,1)}"
                for letter, width in {"A":19.36,"B":19.18,"C":15.63,"D":14.82,"E":14.54,"F":13.0}.items():
                    ws.column_dimensions[letter].width = width
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    header = ws.cell(1, cell.column).value
                    if header in [
                        "Current SS%", "SS% After DC @ Firm PO Week", "SS% Before DC @ Balance Week",
                        "SS% After DC @ Balance Week", "Sub Vendor SS% After", "OSQP SS% After",
                        "OSQP Sub Vendor SS% After", "Priority Rule Value"
                    ] and isinstance(cell.value, (int, float)) and not math.isinf(cell.value):
                        cell.number_format = "0.0%"
    return output_path


# ============================================================
# One-shot process
# ============================================================

def process_files(
    plan_detail_csv: str,
    production_schedule_csv: str,
    due_date_calc_xlsx: str,
    output_path: str,
    target_week: date,
    current_week: Optional[date] = None,
    priority_rules: Optional[Dict[str, PriorityRule]] = None,
    psw_csv_paths: Optional[List[str]] = None,
    due_date_calc_xlsx_list: Optional[List[str]] = None,
    respect_priority_rank: bool = False,
    use_osqp_second_pass: bool = False,
    auto_balance_week: bool = False,
    healthy_ss_pct: float = 150.0,
    freeze_hold_buy: bool = False,
) -> str:
    if current_week is None:
        current_week = saturday_of_current_week()
    if current_week > target_week:
        raise ValueError("Current Week cannot be later than Target Week.")

    due_list = [p for p in (due_date_calc_xlsx_list or []) if p]
    if not due_list:
        due_list = [due_date_calc_xlsx]

    # DueDateCalc mapping follows vendor order detected from PSW.
    offset_map, offset_debug = load_due_date_offsets(due_list[0])
    sub_offset_map = offset_map
    sub_offset_debug = pd.DataFrame(columns=offset_debug.columns)
    if len(due_list) > 1:
        sub_offset_map, sub_offset_debug = load_due_date_offsets(due_list[1])

    all_psw_paths = psw_csv_paths if psw_csv_paths else [production_schedule_csv]
    psw_vendor_df = detect_psw_vendors(all_psw_paths)
    vendor_offset_maps, vendor_due_debug = build_vendor_offset_maps(psw_vendor_df, due_list)

    # Firm PO source / PSW details. First uploaded PSW is the main source by default;
    # within a PSW containing multiple vendors, Timeline-vendor matching identifies the main vendor.
    f_wk3, prod_debug = load_fwk3_from_production(all_psw_paths[0], target_week)
    psw_supply_detail, psw_read_debug = load_psw_vendor_supply(
        all_psw_paths, target_week, current_week, offset_map,
        other_vendor_offset_map=sub_offset_map,
        vendor_offset_maps=vendor_offset_maps,
    )

    # Build the standard firm-week optimizer supply table. This also establishes the reconciliation gap.
    optimizer_base, build_debug, missing_f, plan_offset_debug, psw_vendor_detail, psw_supply_debug = build_optimizer_input_direct_from_plan(
        plan_detail_csv, offset_map, f_wk3, target_week, current_week,
        psw_supply_detail=psw_supply_detail,
        vendor_offset_maps=vendor_offset_maps,
    )

    # Optional automatic balancing week.
    # Firm PO week remains fixed at target_week; balancing metrics are recalculated fresh at each candidate week.
    raw, date_map, first_original_week, last_original_week, first_etd_week = load_plan_source(plan_detail_csv)
    raw_items = raw["Item #"].map(normalize_item).dropna().astype(str).unique().tolist()
    selected_map = {item: target_week for item in raw_items}
    balance_debug = pd.DataFrame([
        [item, "NONE", fmt_date(target_week), float(healthy_ss_pct) / 100.0, "", 0, "", 0, "", 0, None, None, "", "", "", None, "", "Firm week (auto balancing disabled)"]
        for item in raw_items
    ], columns=[
        "Item", "Driver Vendor", "Selected Balance Week", "Healthy SS%",
        "Active WHs", "Active WH Count", "Firm=0 Excluded WHs", "Firm=0 Excluded Count",
        "Active Above Healthy WHs", "Active Above Healthy Count",
        "Min Active SS% After", "Max Active SS% After",
        "Last Positive SS WHs", "Runout WHs", "Future Shortage WHs", "Min Future SI",
        "No Positive SS History WHs", "Reason"
    ])

    if auto_balance_week:
        selected_map, balance_debug = auto_select_balance_week_map(
            raw,
            date_map,
            first_etd_week,
            target_week,
            current_week,
            offset_map,
            optimizer_base,
            priority_rules=priority_rules,
            respect_priority_rank=respect_priority_rank,
            threshold=float(healthy_ss_pct) / 100.0,
            freeze_hold_buy=freeze_hold_buy,
        )

        frames = []
        final_candidate_cache: Dict[date, pd.DataFrame] = {}
        # Per-item source-safe horizon for future-demand protection.
        horizon_src = raw.copy()
        horizon_src["_item"] = horizon_src["Item #"].map(normalize_item)
        horizon_src["_whse"] = horizon_src["Whse"].map(normalize_whse)
        if "MakeBuy Code" in horizon_src.columns:
            horizon_src = horizon_src[horizon_src["MakeBuy Code"].fillna("").astype(str).str.strip().str.upper().eq("B")]

        for balance_week in sorted(set(selected_map.values())):
            candidate_all = _get_balance_candidate_cached(
                final_candidate_cache, raw, date_map, first_etd_week, balance_week, current_week,
                offset_map, optimizer_base, target_week
            )
            items_for_week = [item for item, wk in selected_map.items() if wk == balance_week]
            if candidate_all.empty:
                continue
            enriched_items = []
            for item in items_for_week:
                g = candidate_all[candidate_all["Item"].astype(str).eq(str(item))].copy()
                if g.empty:
                    continue
                main_pool = int(round(pd.to_numeric(g.get("F Wk3", 0.0), errors="coerce").fillna(0.0).sum()))
                sub_pool = int(round(pd.to_numeric(g.get("Other Vendor Supply", 0.0), errors="coerce").fillna(0.0).sum()))
                driver = "MAIN" if main_pool > 0 else ("SUB" if sub_pool > 0 else "NONE")
                whses = horizon_src.loc[horizon_src["_item"].astype(str).eq(str(item)), "_whse"].dropna().astype(str).unique().tolist()
                last_week = _dynamic_balance_candidate_weeks(date_map, target_week, offset_map, whses=whses)[-1]
                future_min_original = _future_min_si_original_map(
                    item, balance_week, last_week, final_candidate_cache, raw, date_map, first_etd_week,
                    current_week, offset_map, optimizer_base, target_week,
                )
                if driver in {"MAIN", "SUB"}:
                    g = _attach_future_firm_floor(g, driver, future_min_original)
                enriched_items.append(g)
            if enriched_items:
                frames.append(pd.concat(enriched_items, ignore_index=True))

        optimizer_input = pd.concat(frames, ignore_index=True) if frames else optimizer_base.copy()
    else:
        optimizer_input = optimizer_base.copy()
        optimizer_input["Balance Week"] = target_week

    detail, summary, detail_full = run_optimizer(
        optimizer_input,
        priority_rules=priority_rules,
        respect_priority_rank=respect_priority_rank,
        freeze_hold_buy=freeze_hold_buy,
    )

    # Output audit layer: keep optimizer internals unchanged, but expose Firm PO Week and Balance Week
    # SI/SS side by side. Firm-week "Current" values are rebuilt fresh from PlanDetailTimeline.
    firm_week_snapshot = _build_balance_candidate_input(
        raw, date_map, first_etd_week, target_week, current_week, offset_map, optimizer_base,
        firm_week=target_week,
    )
    detail_full = attach_output_week_metrics(detail_full, firm_week_snapshot, target_week)
    detail = build_detail_output(detail_full)

    upload_sheet = build_upload_sheet(detail_full, target_week, current_week, psw_vendor_detail=psw_vendor_detail)
    osqp_sheets = build_osqp_sheets(detail_full) if use_osqp_second_pass else {}

    # Add standard audit-only columns required by the approved output.
    run_debug = pd.DataFrame([
        ["Target Week / Firm PO Week", fmt_date(target_week)],
        ["Current Week", fmt_date(current_week)],
        ["DueDateCalc files", ", ".join(due_list)],
        ["PSW files", ", ".join(all_psw_paths)],
        ["DueDateCalc mapping", "Detected PSW vendor order: DueDateCalc #1 -> Vendor #1; #2 -> Vendor #2; missing later files use the last uploaded file as fallback"],
        ["Auto balancing", "Enabled" if auto_balance_week else "Disabled"],
        ["Healthy SS% Threshold", f"{healthy_ss_pct:.1f}%"],
        ["Auto Balance Horizon", "Dynamic through the last fully evaluable PlanDetailTimeline week"],
        ["Last Positive SS", "Bridge only: first zero-SS candidate bucket uses the most recent prior positive SS. From 2 consecutive zero-SS buckets onward, Runout Balance Mode uses future SI instead of SS%."],
        ["Future shortage protection", "Enabled: all change-eligible Buy warehouses are projected through the remaining fully evaluable PlanDetail horizon. The allocator protects a minimum Future Firm Floor and rejects a candidate if projected SI still drops below zero."],
        ["Runout balance", "Enabled: when SS has been zero for 2+ consecutive candidate buckets (or no positive SS history exists), SS% no longer gates the candidate; future SI >= 0 is used."],
        ["Auto balancing rule", "Fixed Healthy SS% only; Reference/Adjusted Healthy removed. Candidate weeks rebuild fresh across all Buy warehouses. Active WH = driver-vendor Firm After > 0 for the Healthy SS% gate; Firm After = 0 is excluded from that SS% gate only for the current candidate, but still participates in future-shortage protection when change-eligible. Main drives when Main Firm > 0; Sub drives when Main Firm = 0 and Sub Firm > 0. No fixed Week-14 cap; the last fully evaluable PlanDetail horizon is used."],
        ["Hold Buy freeze", "Enabled" if freeze_hold_buy else "Disabled"],
        ["UPLOAD sheet", "Firm PO Week is used for WeekIndex; main vendor rows first, then sub-vendor rows; only non-zero destination-change rows are included."],
        ["Priority rank", "Enabled" if respect_priority_rank else "Disabled"],
        ["Hard lock", "Priority SI = 0 keeps original Firm PO quantity fixed"],
        ["Zero-SS fallback", "Automatic SI equalization pass for multiple zero-SS warehouses"],
        ["Total Firm PO Before", int(detail_full["F Wk3 Original"].sum()) if not detail_full.empty else 0],
        ["Total Firm PO After", int(detail_full["F Wk3 After Destination Change"].sum()) if not detail_full.empty else 0],
    ], columns=["Field", "Value"])

    debug_sheets = {
        "Run Debug": run_debug,
        "DueDate Offset Debug": offset_debug,
        "DueDate Sub Offset Debug": sub_offset_debug,
        "Vendor DueDate Mapping": vendor_due_debug,
        "PSW Read Debug": psw_read_debug,
        "PSW Supply Debug": psw_supply_debug,
        "Other Vendor Supply Detail": psw_vendor_detail,
        "Balance Week Debug": balance_debug,
        "Build Debug": build_debug,
        "Plan Offset Debug": plan_offset_debug,
        "Production Debug": prod_debug,
        "Missing F Match": missing_f,
    }
    # Only Optimized Data is part of the normal download; OSQP sheets are optional.
    return write_excel_output(output_path, detail, summary, optimizer_input, debug_sheets, extra_sheets=osqp_sheets, upload_sheet=upload_sheet)


# Backend module ends here.

if __name__ == "__main__":
    raise SystemExit("This module is intended to be imported by the Streamlit application.")
