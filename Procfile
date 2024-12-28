web: gunicorn app:app  # Flask app managing routes
streamlit: streamlit run streamlit_dashboard.py --server.port=$PORT --server.address=0.0.0.0 --server.enableCORS false --server.enableXsrfProtection false
dash: python dash_dashboard.py  # Dash dashboard