# Destination Change v6.10 - Planned Transfer Exclusion

This version keeps all approved v6.9 logic and adds the new Balance-Week planned-transfer rule.

## v6.10 Planned Transfer SI adjustment
From the selected Balance Week forward, projected SI removes the effect of planned transfers already embedded in `SHIPPABLE INV`:

`Adjusted SI = Raw SI - cumulative PLANNED TFR IN + cumulative PLANNED TFR OUT`

Rules:
- `PLANNED TFR IN` is already added into raw SI, so it is subtracted back out.
- `PLANNED TFR OUT` is already subtracted from raw SI, so it is added back.
- The adjustment is cumulative from the selected Balance Week through the evaluated future week.
- Planned transfers before the selected Balance Week remain in the SI baseline.
- The transfer adjustment is applied before water-fill, Healthy SS% evaluation, and the iterative Negative-SI fallback.
- Every fallback week becomes a new Balance Week candidate, so transfer exclusion is rebuilt fresh from that fallback week.
- Optional near-term Promotion / Demand Peak review carries the same selected Balance Week as the transfer-exclusion anchor across its forward stress window.

### Example
If Balance Week is 12/12 and WH17 has:
- Raw SI 12/12 = -55, Planned TFR IN 12/12 = 118 -> Adjusted SI = -173.
- Raw SI 12/19 = -19, Planned TFR IN 12/12 = 118 and 12/19 = 56 -> Adjusted SI = -193.

If Balance Week instead starts at 12/19, the earlier 12/12 transfer remains in the baseline and only the 12/19 transfer is removed.

## Auto Separate / Balance Week logic retained
- Firm PO Week remains Candidate Week 1.
- Each candidate is rebuilt fresh from PlanDetailTimeline.
- Standard PCS-by-PCS water-fill is used for Auto Separate ON and OFF.
- Active WH = driver-vendor `Firm After > 0` and not Hard Lock.
- Candidate passes when every Active WH has `SS% After <= Healthy SS%`; there is no lower bound.
- `Firm After = 0` is excluded only for the current candidate and may become Active again later.
- Main drives when Main Firm > 0; Sub drives when Main Firm = 0 and Sub Firm > 0.

## Iterative Negative-SI fallback retained
If the tentative Balance Week is later than Firm PO Week and any non-hard-locked warehouse has negative final SI after DC:
- move back one candidate week;
- rebuild SI/SS fresh, including the planned-transfer exclusion starting from that fallback week;
- rerun the standard Auto-Separate-OFF water-fill;
- continue week-by-week until no repairable warehouse is negative or Firm PO Week is reached.

Hard Locks such as `Priority SI = 0` do not trigger fallback.

## Output audit retained
`Optimized Data` shows:
- Firm PO Original -> Firm PO After -> Net Destination Change;
- Current SI / SS / SS% and After-DC state at Firm PO Week;
- when Balance Week differs from Firm PO Week, the After-DC SI / Effective SS / SS% at Balance Week.

Firm-Week audit follows the transfer rule precisely:
- if Balance Week = Firm PO Week, Firm-Week SI uses the transfer-adjusted state;
- if Balance Week > Firm PO Week, Firm-Week before-DC SI remains untouched because transfer exclusion starts later.

## Streamlit
Main file: `destination_change_streamlit_app.py`

Install:
`pip install -r requirements.txt`

Run:
`streamlit run destination_change_streamlit_app.py`

OSQP remains optional through `requirements-full.txt` / `requirements-optional-osqp.txt`.
