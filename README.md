# Destination Change v5.6 Stable Streamlit

This package uses the proven v5.0 Streamlit layout and keeps the backend separate.

## Startup
- Streamlit is the only heavy runtime dependency imported during initial page bootstrap.
- Pandas, NumPy, OpenPyXL, and the optimizer backend are loaded only when Run Full Flow is clicked.
- PSW vendor detection uses a lightweight cached CSV parser.

## Run
streamlit run destination_change_streamlit_app.py
