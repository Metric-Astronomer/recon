#!/bin/bash
# Double-click this file to start the reconciliation app.
# It will open automatically in your browser.

cd "$(dirname "$0")"

# Suppress Streamlit's first-run email prompt
mkdir -p ~/.streamlit
echo '[general]' > ~/.streamlit/credentials.toml
echo 'email = ""' >> ~/.streamlit/credentials.toml

python3 -m streamlit run app.py --server.headless true
