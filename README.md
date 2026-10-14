# Destination Change v6.6 - Firm-Week Fresh Water-Fill Fix

This package keeps the unified water-fill logic and locks **Auto separate Firm PO week and balancing week** to the exact same Firm PO redistribution behavior as a standard run. Candidate Week 1 is the Firm PO Week itself, rebuilt from the original Firm-week state, then water-filled PCS-by-PCS before the Healthy gate is evaluated.

## v6.6 Firm-Week fresh allocation fix
- The Main and Sub Firm PO pools are snapshotted once from the selected **Firm PO Week** and kept immutable through candidate-week search.
- Each candidate week rebuilds SI/SS fresh, then runs the same standard water-fill optimizer from that immutable Firm pool.
- `Firm PO Original` always remains the original Firm PO Week quantity.
- `Firm PO After Destination Change` is the selected allocation result only.
- `Net Destination Change = Firm PO After - Firm PO Original`.
- `Current SI / Current SI-SS / Current SS / Current SS%` remain Firm PO Week **before DC** metrics.
- `SI Before DC @ Balance Week` remains the selected candidate baseline before redistribution; `SI After DC @ Balance Week` applies the final Main + Sub DC deltas.
- A fail-fast validation stops the run if an Auto Separate allocation is ever baked back into Main or Sub Firm Original.

Regression lock for item `2140598`:
- Firm PO Week = `10/31/2026`.
- Firm Original remains `33 / 0 / 8 / 10` for WH `15 / 17 / 335 / 5`.
- Firm-week Current SI remains `26 / 10 / 3 / 4`.
- The allocator internally removes the original flexible Firm to form the no-Firm water-fill baseline, then redistributes the same Firm pool one PCS at a time.
- With WH335 hard-locked at 8, the expected Main Firm After is `23 / 0 / 8 / 20`, giving Net DC `-10 / 0 / 0 / +10`.
- Active WH15 = `123.1%` and WH5 = `127.3%` are <= Healthy 150%; WH17 has Firm After = 0 and is excluded for that candidate; WH335 is hard-locked. Therefore Balance Week = `10/31/2026`.

## Streamlit
Main file: `destination_change_streamlit_app.py`

Install default dependencies:
`pip install -r requirements.txt`

Run:
`streamlit run destination_change_streamlit_app.py`

OSQP is optional. Use `requirements-full.txt` only when the OSQP second-pass is needed.

## Scenario priority

### 1) Normal Healthy Range - default
- Firm PO Week remains the user-selected Target Week.
- Reference WH / Adjusted Healthy are removed.
- Healthy SS% is fixed at the user threshold (default 150%).
- Normal candidate weeks are evaluated sequentially through the last fully evaluable PlanDetailTimeline week.
- Every candidate is rebuilt fresh and Destination Change is rerun across all MakeBuy Code = B warehouses.
- **Unified allocation engine:** Auto Separate uses the exact same optimizer pipeline as Auto Separate OFF: Main allocation -> zero-SS fallback -> Sub Vendor allocation. Only the candidate Balance Week changes.
- **Firm Week equivalence lock:** when Candidate Week = Firm PO Week, Auto Separate ON must use the same Firm-week SI/SS/Firm state and return the same warehouse-level water-fill allocation as Auto Separate OFF.
- **Per-PCS water-fill:** each Firm PO unit goes to the warehouse with the lowest current `SS% After`; SS% is recalculated after every unit, so the recipient can switch as warehouses catch up and naturally equalize.
- Hard Locks are excluded from the Balance Week health gate.
- **Active WH = driver-vendor Firm After > 0**.
- **Firm After = 0 is excluded only for that candidate**. The next candidate week is rebuilt fresh, so the warehouse can become Active again if it receives Firm.
- A normal candidate passes when every Active WH is at/below Healthy SS%.
- **There is no lower SS% bound.** Negative, zero, and positive SS% values all pass when `SS% After <= Healthy`.
- Firm PO Week is Candidate Week 1 and uses the exact same rule after Destination Change.
- Main Vendor drives when Main Firm > 0. If Main Firm = 0 and Sub Firm > 0, Sub Vendor uses the same rule from Main Vendor SI After.

Example:

| WH | Firm After | SS% After | Health gate |
|---|---:|---:|---|
| 28 | 0 | 230% | Excluded this candidate |
| 1 | 20 | 125% | Pass |
| 15 | 30 | 140% | Pass |
| 5 | 50 | 165% | Fail -> next candidate |

At the next candidate week WH28 is recalculated from scratch. If its Firm After becomes >0, WH28 becomes Active again and returns to the Normal Healthy Range check.

Example of the clarified no-lower-bound rule: if Healthy = 150%, Active WH values `140%`, `0%`, `-20%`, and `-150%` all pass; `170%` fails.

### 2) Firm After = 0 - exemption inside Normal Healthy Range
This is not a separate permanent status. It only means the warehouse does not block the Healthy SS% gate in the current candidate. There is no permanent participant freeze.

### 3) Near-term Promotion / Demand Peak Review - OPTIONAL
Streamlit includes a separate checkbox:

`Review Near-term Promotion / Demand Peak`

Default = OFF.

- OFF: Balance Week is determined by the Normal Healthy Range only.
- ON: after the normal candidate is allocated with the standard water-fill, the tool also reviews the next 4 weeks for projected SI shortage.
- The promotion review does **not** change Firm allocation or add a pre-allocation Firm floor; it is a post-allocation stress gate only.
- An otherwise Healthy candidate can be rejected only when this optional review is enabled.

This is intended as a separate stress-review mode and does not delay default Balance Week selection.

### 4) Hold Buy / Phase-out - End-of-Timeline Balance
When `Exclude Hold Buy (HB) from destination change` is OFF, an item is routed to End-of-Timeline Balance when:
- it is Hold Buy, or
- it has entered a network-wide phase-out state where all Buy warehouses have SS=0 for at least 2 consecutive final evaluable buckets.

Logic:
- Select the **last fully evaluable PlanDetailTimeline week** as the Balance Week / SI base.
- Rebuild End SI for all Buy warehouses.
- Preserve total Firm PO.
- Keep Hard Locks fixed and enforce FIRM=0 exclusions.
- Allocate available Firm one unit at a time to the eligible warehouse with the **lowest projected End SI**.
- This first minimizes negative End SI exposure and then naturally equalizes the remaining eligible warehouses as evenly as integer Firm quantities allow.

If `Exclude Hold Buy (HB) from destination change` is ON, HB Freeze remains the highest-priority override: HB Firm stays unchanged and End-of-Timeline redistribution is not applied to that HB item.

## Priority / locks
Existing precedence remains:
- HB Freeze when enabled
- `SI=0` hard lock
- `FIRM=0` target
- optional SI / SS priority and rank
- scenario-specific balancing logic

WH335 is not automatically fixed; it is fixed only when the user setup makes it a hard lock, for example `SI=0`.

## Optimized Data columns
The v6.1 user-facing column structure is retained.

Firm PO Week before DC:
- `Current SI`
- `Current SI-SS`
- `Current SS`
- `Current SS%`

Firm PO Week after DC:
- `SI After DC @ Firm PO Week`
- `SI-SS After DC @ Firm PO Week`
- `SS% After DC @ Firm PO Week`

Balance Week:
- `SI Before DC @ Balance Week`
- `SI-SS Before DC @ Balance Week`
- `Effective SS @ Balance Week`
- `SS% Before DC @ Balance Week`
- `SI After DC @ Balance Week`
- `SI-SS After DC @ Balance Week`
- `SS% After DC @ Balance Week`

`Firm PO Week` remains next to `Balance Week` near the end of the output. Priority, reconciliation, Sub Vendor, Vendor, and metadata blocks retain their previous relative order.

## UPLOAD
- Firm PO Week is used for WeekIndex.
- Main Vendor rows first, then Sub Vendor rows.
- Only non-zero destination-change rows are included.
- Quantity is the After quantity.
- A donor whose After becomes 0 is retained with `Quantity = 0` so the previous Firm PO can be cleared.
