# Destination Change v5.4 - Fast Boot Streamlit

Main Streamlit entrypoint: `destination_change_streamlit_app.py`
Alternative entrypoint: `app.py`

The UI starts with only Streamlit imported. The optimizer backend and heavy libraries are lazy-loaded only when Run Full Flow is pressed. The package is flat (no nested folder), which avoids path problems when deploying the main file directly to Streamlit Cloud or running locally.

Business logic is preserved from v5.3.
