web: gunicorn app:app  # Flask app managing routes
worker: streamlit run streamlit_dashboard.py --server.port $PORT --server.enableCORS false  # Streamlit dashboard
dash: python dash_dashboard.py  # Dash dashboard