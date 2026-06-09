import os
import jwt
import datetime
from flask import Flask, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

JWT_SECRET = os.environ.get("JWT_SECRET")
SUPABASE_STAFF_URL = os.environ.get("SUPABASE_STAFF_URL")
SUPABASE_STAFF_KEY = os.environ.get("SUPABASE_STAFF_KEY")

@app.route("/health_staff")
def health():
    return jsonify({"status": "ok", "service": "staff-dashboard"})

if __name__ == "__main__":
    app.run(debug=False)
