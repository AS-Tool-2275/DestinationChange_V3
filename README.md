# Destination Change v5.2

## UPLOAD sheet donor-zero fix
- Uses Firm PO Week for WeekIndex.
- Main vendor rows are written first, followed by sub-vendor rows.
- Includes every row with non-zero destination change, including negative donor rows whose Firm PO After Destination Change equals 0.
- Quantity is always the Firm PO After Destination Change value, so Quantity = 0 is intentionally retained to clear the donor warehouse/vendor in the target system.
- Sub-vendor zero rows are retained for vendors with source quantities at Firm PO Week when the sub-vendor destination change is non-zero.
