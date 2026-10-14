# Destination Change v6.7 - Simplified Before / After Audit

This version keeps the v6.6 unified PCS-by-PCS water-fill and Auto Separate candidate logic, while simplifying and hardening the user-facing `Optimized Data` audit view.

## Allocation and Balance Week logic
- Firm PO Week remains Candidate Week 1.
- Every candidate week is rebuilt fresh from PlanDetailTimeline.
- Auto Separate ON uses the same standard allocation pipeline as Auto Separate OFF.
- Firm is water-filled one PCS at a time to the currently lowest `SS% After`; the recipient is recalculated after every PCS.
- Active WH = driver-vendor `Firm After > 0` and not Hard Lock.
- Candidate passes when every Active WH is `SS% After <= Healthy SS%`.
- There is no lower SS% bound; negative SS% can pass.
- `Firm After = 0` is excluded only for the current candidate and may become Active again at the next candidate.
- Main drives when Main Firm > 0. If Main Firm = 0 and Sub Firm > 0, Sub uses the same candidate logic with Main SI After as its baseline.

## v6.7 audit fix
The Firm PO Week snapshot is authoritative for all user-facing **before DC** values.

For every Item + Warehouse:
- `Firm PO Original` = original Main-vendor Firm quantity at Firm PO Week.
- `Firm PO After Destination Change` = final selected allocation.
- `Net Destination Change = After - Original`.
- `Current SI`, `Current SS`, `Current SS%` = Firm PO Week **before DC**.
- `SI After DC @ Firm PO Week`, `SS% After DC @ Firm PO Week` = Firm PO Week **after the final Main + Sub DC deltas**.

If `Balance Week != Firm PO Week`, the output also shows:
- `SI After DC @ Balance Week`
- `Effective SS @ Balance Week`
- `SS% After DC @ Balance Week`

If `Balance Week == Firm PO Week`, those three Balance-Week columns are intentionally blank because the Firm-Week After block already contains the same final state.

The normal `Optimized Data` sheet no longer displays the intermediate Balance-Week-before-DC and SI-SS audit columns. They remain internal to the backend where needed.

## Regression examples locked
### Healthy 150% - item 2140598
Firm PO Week = `10/31/2026`, original Main Firm by WH 15 / 17 / 335 / 5 = `33 / 0 / 8 / 10`.

Expected Firm-Week water-fill:
- Firm After = `23 / 0 / 8 / 20`
- Net DC = `-10 / 0 / 0 / +10`
- Firm-Week Current SI before DC = `26 / 10 / 3 / 4`
- Firm-Week SI After DC = `16 / 10 / 3 / 14`
- WH15 = 123.1%, WH5 = 127.3%; WH17 Firm After = 0 is excluded; WH335 is Hard Lock.
- Balance Week = `10/31/2026`.
- Balance-Week-after display columns are blank because Balance Week equals Firm PO Week.

### Healthy 100% - item 2140598
The 10/31 candidate fails because WH15 and WH5 are above 100%. If the selected Balance Week is `11/07/2026`, the user-facing audit must still retain the original Firm Week values:
- Firm Original = `33 / 0 / 8 / 10`
- Final Firm After = `0 / 43 / 8 / 0`
- Net DC = `-33 / +43 / 0 / -10`
- Current SI @ Firm Week before DC = `26 / 10 / 3 / 4`
- SI After DC @ Firm Week = `-7 / 53 / 3 / -6`
- SI After DC @ Balance Week = `-11 / -35 / 3 / -9`

This separates the original Firm-Week state from the later selected candidate state and prevents the selected allocation from being displayed as the original Firm PO.

## Streamlit
Main file: `destination_change_streamlit_app.py`

Install:
`pip install -r requirements.txt`

Run:
`streamlit run destination_change_streamlit_app.py`

OSQP remains optional through `requirements-full.txt` / `requirements-optional-osqp.txt`.

## Existing controls retained
- Optional Healthy SS% threshold, default 150%.
- Optional `Review Near-term Promotion / Demand Peak`: post-allocation 4-week stress gate; it does not alter the standard water-fill allocation.
- Optional Hold Buy freeze.
- HB / phase-out End-of-Timeline logic when HB Freeze is OFF.
- SI=0 hard lock.
- FIRM=0 exclusion.
- Optional SI / SS priority and Priority Rank.
- Zero-SS fallback equalization.
- Vendor-aware Main/Sub flow and DueDateCalc mapping.
- UPLOAD sheet uses Firm PO Week for WeekIndex and retains zero-quantity donor rows when needed to clear prior Firm PO.

## Optimized Data visible columns
The simplified sheet keeps the before/after planning view first, followed by priority, reconciliation, Sub Vendor, vendor/week, and the nine PlanDetail metadata columns.

Primary audit block:
1. Item
2. ProdResourceID
3. Whse
4. Firm PO Original
5. Firm PO After Destination Change
6. Net Destination Change
7. Firm PO Total
8. Current SI
9. Current SS
10. Current SS%
11. SI After DC @ Firm PO Week
12. SS% After DC @ Firm PO Week
13. SI After DC @ Balance Week - blank when same as Firm PO Week
14. Effective SS @ Balance Week - blank when same as Firm PO Week
15. SS% After DC @ Balance Week - blank when same as Firm PO Week
