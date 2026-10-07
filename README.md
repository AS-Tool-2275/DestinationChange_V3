# Destination Change v5.7 - Stable Streamlit

This package keeps the proven v5.0 Streamlit UI structure and updates the backend Auto Separate logic.

## Streamlit
Main file: `destination_change_streamlit_app.py`

Install default dependencies:
`pip install -r requirements.txt`

Run:
`streamlit run destination_change_streamlit_app.py`

OSQP is optional. Use `requirements-full.txt` only when the OSQP second-pass is needed.

## v5.7 Auto Separate logic
- Firm PO Week remains the user-selected Target Week.
- Candidate Balance Weeks are Week 1 through Week 14, where Week 1 = Firm PO Week.
- At Week 1, the participant set is frozen as Receiver + Giver only (`Net Destination Change != 0`).
- Neutral warehouses (`Net Destination Change = 0`) do not participate by default.
- If every Week-1 Receiver/Giver is above the user Healthy SS% threshold, exactly one eligible neutral reference warehouse is added: the neutral warehouse with the highest finite Week-1 SS% above the default threshold.
- The reference warehouse Week-1 SS% becomes the adjusted Healthy SS% threshold for that item.
- Reference warehouse and adjusted threshold are determined once at Week 1 and remain fixed through Week 14.
- Main Vendor drives Auto Separate when Main has Week-1 destination change. If Main has no Week-1 destination change, Sub Vendor performs its own Auto Separate evaluation using Main Vendor SI After as its baseline.
- WH335 is not automatically excluded. It is excluded from reference eligibility only when user setup makes it a hard lock, such as Priority `SI = 0`.
- HB rows are excluded only when `Exclude Hold Buy (HB) from destination change` is enabled.
- Priority `FIRM = 0` forces Firm PO After = 0 and is validated after optimization.

## SI / SS basis
At Balance Week, SI subtracts Planned POS through Balance Week and all Firm POS from Firm PO Week through Balance Week, inclusive. Supply is added exactly once. Main SS% After uses Main SI After / SS at Balance Week. Sub Vendor SS% After starts from Main SI After and applies Sub Vendor Net Destination Change.

## UPLOAD
- Uses Firm PO Week for WeekIndex.
- Main Vendor rows first, then Sub Vendor rows.
- Includes only rows with non-zero destination change.
- Quantity is the After quantity.
- A donor whose After quantity becomes 0 is retained as `Quantity = 0` so the prior Firm PO can be cleared.
