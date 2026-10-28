# Destination Change v6.12 - Protected Path + Warehouse Scope

This package is based on v6.11 and keeps the existing Main/Sub Vendor, Healthy SS%, Planned Transfer exclusion, Priority, Hard Lock, FIRM=0, HB, zero-SS and audit logic.

## v6.12 changes

- Default Destination Change warehouse scope: `1, 15, 17, 28, 335, 5, ECR`. Warehouses outside this list do not participate in Firm pools, Healthy gates, Balance Week selection, or water-fill.
- Protected-path water-fill for Auto Separate:
  - Find the tentative Balance Week using the existing v6.11 Healthy logic.
  - A negative SI at the Balance Week itself does **not** force fallback.
  - For a tentative Balance Week `W`, every non-Hard-Lock warehouse must remain `SI >= 0` from Firm PO Week through `W-1`.
  - The engine reserves the minimum Firm needed to protect that path, then allocates all remaining Firm using the existing PCS-by-PCS lowest SS% water-fill at `W`.
  - `W+1` is used only as a directional tie-break when Balance-Week SS% is tied.
  - `W-1` is a soft Healthy-range tie-break; it is not required to be <= Healthy.
  - If the protected path is infeasible because of priorities/locks/insufficient Firm, fallback proceeds one week at a time and never before Firm PO Week.
- Main/Sub parity is retained. When Main Firm = 0 and Sub Firm > 0, Sub remains the Auto Balance Driver and uses the same protected-path logic.
- Planned Transfer rule from v6.10/v6.11 is retained: from Balance Week forward, SI removes cumulative `PLANNED TFR IN` and restores cumulative `PLANNED TFR OUT`. Transfers before Balance Week remain embedded.

## Regression case 7471494

With Firm PO Week 11/14/2026, Healthy 150%, WH335 Priority SI=0, and the supplied Timeline/PSW/DueDateCalc inputs, v6.12 selects Balance Week 11/21/2026. Sub Vendor Firm remains 333 at WH17 and 0 at WH5 because WH5 stays non-negative on the protected path; WH17 being negative at 11/21 no longer forces a fallback to 11/14.

## Existing package notes

# Destination Change v6.11 - Sub Vendor Auto-Separate Parity

This version keeps all approved v6.10 logic, including Planned Transfer exclusion and iterative Negative-SI fallback, and makes the Sub Vendor allocation pipeline fully parallel to Main Vendor when Auto Separate is used.

## v6.11 Sub Vendor parity
When Main Firm PO = 0 and Sub Vendor Firm PO > 0, Sub Vendor is the `Auto Balance Driver`.

Sub Vendor now uses the same planning pipeline as Main Vendor:
- same PCS-by-PCS lowest-SS% water-fill (`allocate_item()`);
- same Priority / Priority Rank behavior;
- same `SI=0` Hard Lock behavior;
- same `FIRM=0` exclusion;
- same Hold Buy freeze behavior;
- same END_TIMELINE mode for Hold Buy / phase-out;
- same zero-SS equalization post-pass;
- same Healthy SS% candidate gate;
- same candidate-by-candidate fresh rebuild;
- same Planned TFR IN / OUT SI adjustment from Balance Week forward;
- same iterative Negative-SI fallback.

Sub Vendor still starts from the completed Main Vendor result:
`Sub Vendor SI Before = Main Vendor SI After`.

## Sub Vendor role label
`Sub Vendor DC Note` is now item-level:
- `Auto Balance Driver` when Main Firm = 0 and Sub Firm > 0;
- `Suggestion only` when Main Firm > 0 and Sub Firm > 0;
- blank when there is no Sub Firm supply.

The label is shown on all warehouses for the item, including recipient warehouses whose original Sub Firm was zero.

## Auto Separate / Balance Week logic retained
- Firm PO Week remains Candidate Week 1.
- Each candidate is rebuilt fresh from PlanDetailTimeline.
- Standard PCS-by-PCS water-fill is used for Auto Separate ON and OFF.
- Active WH = driver-vendor `Firm After > 0` and not Hard Lock.
- Candidate passes when every Active WH has `SS% After <= Healthy SS%`; there is no lower bound.
- `Firm After = 0` is excluded only for the current candidate and may become Active again later.
- Main drives when Main Firm > 0; Sub drives when Main Firm = 0 and Sub Firm > 0.
- If the PlanDetail horizon is reached before Healthy is achieved, the last fully evaluable week is retained, same for Main and Sub.

## Planned Transfer SI adjustment retained
From the selected Balance Week forward:

`Adjusted SI = Existing SI - cumulative PLANNED TFR IN + cumulative PLANNED TFR OUT`

Transfers before Balance Week remain in the SI baseline. Every fallback candidate rebuilds this adjustment from its own Balance Week.

## Iterative Negative-SI fallback retained
If tentative Balance Week > Firm PO Week and any non-hard-locked warehouse has negative final SI after Main + Sub DC, move backward one candidate week at a time, rerun the standard allocator, and stop at the first non-negative week or at Firm PO Week.

Hard Locks such as `Priority SI = 0` do not trigger fallback.

## Output audit retained
`Optimized Data` shows:
- Firm PO Original -> Firm PO After -> Net Destination Change;
- Current SI / SS / SS% and After-DC state at Firm PO Week;
- when Balance Week differs from Firm PO Week, After-DC SI / Effective SS / SS% at Balance Week;
- Sub Vendor Original / After / Net DC / SI Before / SI After / SS% After / role note.

## Streamlit
Main file: `destination_change_streamlit_app.py`

Install:
`pip install -r requirements.txt`

Run:
`streamlit run destination_change_streamlit_app.py`

OSQP remains optional through `requirements-full.txt` / `requirements-optional-osqp.txt`.
