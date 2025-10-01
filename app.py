# Flask app code with Basic Auth, Missing Tickets, Summary sheet
import os
from functools import wraps
from flask import Flask, render_template, request, send_file, flash, redirect, url_for, Response
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import tempfile

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret-change-me")

# --- Basic Auth (set APP_USER and APP_PASS in Render env vars) ---
USERNAME = os.getenv("APP_USER")
PASSWORD = os.getenv("APP_PASS")

def requires_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not USERNAME or not PASSWORD:  # no creds set -> skip auth
            return f(*args, **kwargs)
        auth = request.authorization
        if not auth or auth.username != USERNAME or auth.password != PASSWORD:
            return Response("Login required", 401, {"WWW-Authenticate": 'Basic realm="Login"'})
        return f(*args, **kwargs)
    return decorated

ALLOWED_EXTENSIONS = {"csv", "xlsx", "xls"}
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def read_table(file_storage):
    name = file_storage.filename.lower()
    if name.endswith(".csv"):
        return pd.read_csv(file_storage)
    else:
        return pd.read_excel(file_storage