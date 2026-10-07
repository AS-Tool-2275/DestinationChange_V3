# Destination Change v6.0 - Hybrid Healthy / Future Shortage / Runout Balance

This package keeps the proven v5.0-style Streamlit UI and updates **Auto separate Firm PO week and balancing week** with the approved Hybrid Balance logic.

## Streamlit
Main file: `destination_change_streamlit_app.py`

Install default dependencies:
`pip install -r requirements.txt`

Run:
`streamlit run destination_change_streamlit_app.py`

OSQP is optional. Use `requirements-full.txt` only when the OSQP second-pass is needed.

## v6.0 Auto Separate logic
- Firm PO Week remains the user-selected Target Week.
- Reference WH / Adjusted Healthy are removed completely.
- Healthy SS% is fixed at the user threshold (default 150%, user adjustable).
- Candidate Balance Weeks are evaluated sequentially from Firm PO Week through the **last fully evaluable PlanDetailTimeline week** for the item; there is no fixed Week-14 cap.
- Each candidate is rebuilt fresh from PlanDetailTimeline and Destination Change is rerun across all MakeBuy Code = B warehouses.
- Main Vendor drives when Main Firm > 0. If Main Firm = 0 and Sub Firm > 0, Sub Vendor becomes the Auto Separate driver and uses the same Hybrid logic starting from Main Vendor SI After.

## 1) Healthy SS% - normal planning
For the driver vendor:
- **Active WH = Firm After Destination Change > 0**.
- A normal-planning Active WH must be at/below the fixed Healthy SS%.
- **Firm After = 0 is excluded from the Healthy SS% gate for that candidate only**, even if its SS% is high.
- The exclusion is not permanent: the warehouse is recalculated from scratch next candidate week and can become Active again.

## 2) Future shortage protection - demand peak guard
The candidate also scans all **change-eligible Buy warehouses** from the candidate week through the remaining fully evaluable PlanDetail horizon.

For each warehouse, the tool calculates the minimum projected SI under the original Firm distribution. It converts that into a minimum safe driver-vendor Firm quantity:

`Future Firm Floor = max(0, CEILING(Original Driver Firm - Minimum Future SI under original distribution))`

Before normal SS balancing, the allocator protects this floor for non-priority/change-eligible warehouses. After redistribution, the candidate is rejected if any change-eligible Buy warehouse still has projected future SI below 0.

This guard includes a warehouse even when its current `Firm After = 0`, so a Neutral / current-zero-Firm warehouse with an upcoming demand peak cannot be ignored. User hard locks and `FIRM=0` targets are excluded because the optimizer is not allowed to repair them.

## 3) Runout balance - when SS loses meaning
- If candidate Safety Stock is positive, use it normally.
- At the **first zero-SS candidate bucket**, use the most recent prior positive SS as a Last Positive SS bridge.
- From **2 consecutive zero-SS candidate buckets onward**, the warehouse enters **Runout Balance Mode**.
- In Runout Mode, SS% no longer decides the candidate. The governing condition is projected future SI >= 0 through the remaining fully evaluable horizon.
- If no prior positive SS exists, the row is treated as Runout immediately.

This prevents phase-out items from being forced through meaningless SS% comparisons while still protecting the remaining demand.

## Combined candidate rule
A candidate week passes when:
1. Every Active normal-planning WH is within Healthy SS%.
2. No change-eligible Buy WH has projected future SI below 0.
3. Active Runout WHs are future-safe (projected minimum SI >= 0).

If the candidate fails, the tool evaluates the next week. If no candidate passes, the last fully evaluable PlanDetail week is used as the source-horizon fallback.

## Active Firm example

| WH | Firm After | SS% After | Healthy SS% gate |
|---|---:|---:|---|
| 1 | >0 | 125% | Pass |
| 15 | >0 | 140% | Pass |
| 17 | >0 | 135% | Pass |
| 28 | 0 | 230% | Excluded for this candidate |
| 5 | >0 | 145% | Pass |
| ECR | >0 | 120% | Pass |

WH28 does not block the Healthy SS% gate because its Firm After is 0, but if WH28 is change-eligible and a later demand peak would make its projected SI negative, the separate Future Shortage Protection can still reject the candidate.

## Priority / locks
Existing precedence remains:
- HB freeze (when enabled)
- `SI=0` hard lock
- `FIRM=0` target
- optional SI / SS priority and rank
- future-demand floor for normal non-priority allocation
- standard lowest-SS% balancing

WH335 is not automatically excluded; it is fixed only when the user setup makes it a hard lock, for example `SI=0`.

## Output / UPLOAD
- Firm PO Week is used for UPLOAD WeekIndex.
- Main Vendor rows first, then Sub Vendor rows.
- Only non-zero destination-change rows are included.
- Quantity is the After quantity.
- A donor whose After quantity becomes 0 is retained as `Quantity = 0` so the prior Firm PO can be cleared.
- There is no separate output-level INF / -INF sanitization rule in v6.0; zero-SS behavior is handled by the planning logic above.
