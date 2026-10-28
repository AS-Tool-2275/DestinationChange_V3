# Destination Change v6.13 - Sequential Firm Waterfall

v6.13 replaces the normal Auto Separate Healthy-threshold candidate scan with a sequential Firm waterfall. The standard single-week optimizer remains available when Auto Separate is OFF.

## Core rule

The Auto Separate target is fixed at **100% SS coverage**:

`SI = Safety Stock`  /  `SI-SS = 0`

There is no Healthy % input in the Streamlit UI.

For each normal item:

1. Start at Firm PO Week with the Firm pool from the selected PSW week.
2. For each warehouse in scope, calculate the Firm needed to reach 100% coverage.
3. If the pool can fully cover the week, allocate only that need and carry the remaining Firm to the next week.
4. Repeat sequentially until the first week where the remaining Firm cannot bring all eligible warehouses to 100%.
5. That week becomes Balance Week.
6. At Balance Week, allocate every remaining PCS to the current lowest SS% warehouse, recalculating after every PCS, so final SS% converges as evenly as integer quantities allow.
7. If every evaluable week is fully covered, the last evaluable week becomes Balance Week and excess Firm is balanced there.
8. Final Firm After by warehouse is the accumulated allocation across the waterfall. Total Firm is preserved by item.

## Planned Transfer rule

Pre-Balance weeks retain the transfer effects already embedded in SHIPPABLE INV.

From the selected Balance Week forward:

`Adjusted SI = Existing SI - cumulative PLANNED TFR IN + cumulative PLANNED TFR OUT`

## Main / Sub Vendor

- Main Vendor runs first when Main Firm > 0.
- Sub Vendor uses the same Sequential Firm Waterfall from Main Vendor SI After.
- If Main Firm = 0 and Sub Firm > 0, Sub Vendor becomes the driver.
- When Main is the driver, Sub uses the same selected Balance Week so the item has one coherent Balance Week.

## Existing rules retained

- Destination Change warehouse scope: `1, 15, 17, 28, 335, 5, ECR`.
- Priority SI=0 = Hard Lock; original Firm is preserved.
- FIRM=0 cannot receive Firm.
- Optional Hold Buy freeze is retained.
- Hold Buy / sustained network phase-out keeps the existing End-of-Timeline logic when HB Freeze is OFF.
- Firm Original audit fields remain immutable.
- UPLOAD sheet format and Firm conservation checks are retained.

## Performance architecture

The new Auto Separate flow scans the timeline forward once per item and stops at the first constrained week. Candidate-week calculations are cached and reused. It no longer performs the previous Healthy candidate scan plus repeated protected-path fallback checks.

## Regression example: 7471494

Using the supplied Timeline / PSW / DueDateCalc and WH335 Priority SI=0, the sequential waterfall fully protects the earlier weeks and selects the first constrained week based on the new 100% rule. Because Planned Transfers before Balance Week remain embedded, the selected Balance Week can be later than v6.12. At the final Balance Week, the non-hard-locked warehouses are balanced by lowest SS% water-fill.

## Run

Install:

`pip install -r requirements.txt`

Run Streamlit:

`streamlit run destination_change_streamlit_app.py`

OSQP remains optional through `requirements-full.txt` / `requirements-optional-osqp.txt`.
