web: gunicorn app.flask_app:app --bind 0.0.0.0:$PORT --workers 3
streamlit: streamlit run streamlit_dashboard.py --server.port=8501 --server.address=0.0.0.0