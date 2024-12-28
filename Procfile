web: gunicorn app:app  # Flask app managing routes
streamlit: streamlit run streamlit_dashboard.py --server.port $PORT --server.enableCORS false
dash: python dash_dashboard.py  # Dash dashboard