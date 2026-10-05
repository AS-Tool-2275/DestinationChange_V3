# Destination Change v5.3 - Fast Streamlit UI

- Fast initial page load: backend optimizer is lazy-loaded only when Run Full Flow is pressed.
- No pandas/numpy/openpyxl import during initial Streamlit page rendering.
- Vendor detection uses lightweight CSV parsing and session-state caching.
- Backend business logic is preserved from v5.2, including multi-vendor, Auto Separate, HB freeze, FIRM=0, 14-week cap, and UPLOAD donor-zero behavior.
