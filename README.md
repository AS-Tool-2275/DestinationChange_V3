# Destination Change v4.7 Optimized

This package keeps the v4.6 business logic and optimizes the Streamlit startup and rerun path.

## Main improvements
- Uses the actual Streamlit UI entry point.
- Backend optimizer is lazy-loaded, so the initial page does not import the full optimizer stack.
- PSW vendor detection is cached across Streamlit reruns.
- Multiple DueDateCalc files remain supported for vendor-specific transit mapping.
- OSQP/Scipy are optional; the default Streamlit environment does not need them for startup.
- No __pycache__ files are shipped.

## Run
1. Install `requirements.txt`.
2. Optional: install `requirements-optional-osqp.txt` to enable the OSQP second-pass.
3. Run `streamlit run destination_change_streamlit_app.py`.

## Business logic
The optimizer logic is carried forward from v4.6, including Auto Separate Firm PO Week vs Balancing Week, Healthy SS% checking, multi-vendor handling, and the v4.6 SI calculation changes.


## Auto Separate safety cap
- Maximum Balance Week evaluation horizon is 14 candidate weeks, with Firm PO Week counted as Week 1.
- If a candidate beyond Week 14 would otherwise be selected, Week 14 is used.
- This prevents out-of-range / zero-base SS from driving %SS selection.


## Latest additions
- Optional Hold Buy (HB) freeze: when enabled, HB Item + Warehouse rows keep original Firm PO and receive no destination change.
- UPLOAD sheet follows the supplied WN3 template columns: Vendor, ProductionResource, Item, WarehouseCode, WeekIndex, Quantity.
- UPLOAD uses Firm PO Week for WeekIndex, lists main-vendor rows first, then sub-vendor rows, and includes only rows with non-zero destination change.
- Sub-vendor upload rows use Other Vendor List; when multiple sub-vendors exist, the post-change quantity is split using the original Firm-Week vendor share.

## v4.9 change summary
1. Optional Hold Buy (HB) freeze. When enabled, HB Item + Warehouse rows keep original Firm PO and are excluded from both main and sub-vendor destination change.
2. UPLOAD sheet follows WN3.xlsx: Vendor, ProductionResource, Item, WarehouseCode, WeekIndex, Quantity.
3. UPLOAD uses selected Firm PO Week for WeekIndex; main-vendor rows are written first, followed by sub-vendor rows.
4. Only rows with non-zero destination change are included in UPLOAD; Quantity is the corresponding Firm PO after destination change (main) or Sub Vendor Firm PO after destination change (sub). Zero Quantity is retained when the destination change itself is non-zero.
5. Sub-vendor upload Vendor is taken from Other Vendor List. Multiple sub-vendors are split by their original Firm-Week PSW quantity share.
