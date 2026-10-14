# Destination Change v6.9 - Iterative Negative-SI Fallback

This version keeps the unified PCS-by-PCS water-fill and simplified before/after audit view from v6.8, and extends the Auto Separate negative-SI crosscheck from a single fallback week to an iterative week-by-week fallback.

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

## v6.9 iterative negative-SI crosscheck
After Auto Separate finds a tentative Balance Week:
- If `Tentative Balance Week == Firm PO Week`, keep the normal result.
- If `Tentative Balance Week > Firm PO Week`, inspect **final SI after Destination Change** across all non-hard-locked warehouses, not only Active WHs.
- If every repairable warehouse has `SI After DC >= 0`, keep the tentative Balance Week.
- If any repairable warehouse has `SI After DC < 0`, move back one candidate week and run the exact standard Auto-Separate-OFF water-fill at that earlier week.
- Recheck final SI after DC at the fallback week.
- If a repairable warehouse is still negative, move back another week and repeat.
- Stop at the **first earlier week with no repairable negative SI**, or at **Firm PO Week** if negative SI still remains there.
- The fallback path does not need to pass the Healthy SS% gate again. Its purpose is to prevent selecting a later Balance Week after the network has already entered negative SI.
- Hard Locks, including SI=0 / HB-freeze locks, do not trigger this fallback because the optimizer cannot repair them.

### Example fallback path
If the tentative Healthy week is 11/21:

`11/21 negative -> 11/14 negative -> 11/07 negative -> 10/31 non-negative`

then Final Balance Week = `10/31`, and the final allocation is the standard Auto-Separate-OFF water-fill calculated at 10/31.

If 10/31 is the Firm PO Week and it is still negative, stop at 10/31. Auto Separate never selects a week before Firm PO Week.

## Simplified before / after audit
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

## Regression example - item 2140598
Firm PO Week = `10/31/2026`, original Main Firm by WH 15 / 17 / 335 / 5 = `33 / 0 / 8 / 10`.

Standard Firm-Week water-fill:
- Firm After = `23 / 0 / 8 / 20`
- Net DC = `-10 / 0 / 0 / +10`
- Firm-Week Current SI before DC = `26 / 10 / 3 / 4`
- Firm-Week SI After DC = `16 / 10 / 3 / 14`

Healthy 150%:
- WH15 = 123.1%, WH5 = 127.3%.
- WH17 Firm After = 0 is excluded; WH335 is Hard Lock.
- Balance Week = `10/31/2026`.

Healthy 100%:
- 10/31 fails the Healthy gate.
- 11/07 can become the tentative Healthy candidate, but repairable warehouses are negative after DC.
- The negative-SI crosscheck falls back week-by-week. In this case it reaches 10/31 and uses the standard 10/31 water-fill above.

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
