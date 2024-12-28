#!/bin/bash
# Start Flask app in the background
gunicorn app.flask_app:app --bind 0.0.0.0:$PORT &
# Start Streamlit app
streamlit run streamlit_dashboard.py --server.port=8501 --server.address=0.0.0.0