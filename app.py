from flask import Flask, render_template, request, redirect, url_for, session
from sqlalchemy import create_engine, text
import os

app = Flask(__name__)
app.secret_key = "your_secret_key"

# Database configuration
DATABASE_URL = os.getenv("DATABASE_URL")
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)
engine = create_engine(DATABASE_URL)

# Flask route for login
@app.route("/", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]

        # Query the database for the user
        with engine.connect() as conn:
            query = text("SELECT * FROM users WHERE username = :username AND password = :password")
            result = conn.execute(query, {"username": username, "password": password}).fetchone()

        # Validate user
        if result:
            session["logged_in"] = True
            session["username"] = username
            return redirect(url_for("menu"))
        else:
            return render_template("login.html", error="Invalid credentials")

    return render_template("login.html")


@app.route("/menu")
def menu():
    if not session.get("logged_in"):
        return redirect(url_for("login"))
    return render_template("menu.html")


@app.route("/logout")
def logout():
    session.pop("logged_in", None)
    session.pop("username", None)
    return redirect(url_for("login"))


@app.route("/dash")
def dash_dashboard():
    if not session.get("logged_in"):
        return redirect(url_for("login"))
    os.system("python dash_dashboard.py")  # Launch Dash
    return "Dash dashboard is running!"


@app.route("/streamlit")
def streamlit_dashboard():
    if not session.get("logged_in"):
        return redirect(url_for("login"))
    # Redirect to the Streamlit dyno URL
    return redirect("https://sales-stock-dashboard-8395a052347b.herokuapp.com/streamlit")


if __name__ == "__main__":
    app.run(debug=True)