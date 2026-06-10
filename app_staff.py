import os
import jwt
import datetime
import bcrypt
from flask import jsonify, request
from supabase import create_client
from openpyxl import load_workbook
from functools import wraps

SUPABASE_STAFF_URL = os.environ.get("SUPABASE_STAFF_URL")
SUPABASE_STAFF_KEY = os.environ.get("SUPABASE_STAFF_KEY")
JWT_SECRET = os.environ.get("JWT_SECRET")

supabase_staff = create_client(SUPABASE_STAFF_URL, SUPABASE_STAFF_KEY)

MASTER_PATH = os.path.join(os.path.dirname(__file__), "スタッフマスター.xlsx")

B_TO_D = {
    "B0000106": "D0000295",
    "B0000107": "D0000326",
    "D0001318": "B0000095"
}

RATE_TABLE = {
    5: {22: (2.30, 2.00), 21: (2.30, 2.00), 20: (2.45, 2.10),
        19: (2.45, 2.10), 18: (2.45, 2.10), 17: (2.60, 2.20),
        16: (2.60, 2.20), 15: (2.60, 2.20), 14: (2.60, 2.20),
        13: (2.75, 2.30), 12: (2.75, 2.30), 11: (2.75, 2.30),
        10: (2.75, 2.30), 9: (2.75, 2.30), 8: (2.95, 2.57),
        7: (2.95, 2.57), 6: (2.95, 2.57), 5: (2.95, 2.57)},
    4: {22: (2.30, 2.00), 21: (2.30, 2.00), 20: (2.45, 2.10),
        19: (2.45, 2.10), 18: (2.45, 2.10), 17: (2.60, 2.20),
        16: (2.60, 2.20), 15: (2.60, 2.20), 14: (2.60, 2.20),
        13: (2.75, 2.30), 12: (2.75, 2.30), 11: (2.75, 2.30),
        10: (2.75, 2.30), 9: (2.75, 2.30), 8: (2.95, 2.57),
        7: (2.95, 2.57), 6: (2.95, 2.57), 5: (2.95, 2.57)},
    3: {22: (2.30, 2.00), 21: (2.30, 2.00), 20: (2.45, 2.10),
        19: (2.45, 2.10), 18: (2.45, 2.10), 17: (2.60, 2.20),
        16: (2.60, 2.20), 15: (2.60, 2.20), 14: (2.60, 2.20),
        13: (2.75, 2.30), 12: (2.75, 2.30), 11: (2.75, 2.30),
        10: (2.75, 2.30), 9: (2.75, 2.30), 8: (2.95, 2.57),
        7: (2.95, 2.57), 6: (2.95, 2.57), 5: (2.95, 2.57)},
}

def load_staff_master():
    wb = load_workbook(MASTER_PATH, data_only=True)
    ws1 = wb["スタッフマスター"]
    ws2 = wb["時給マスター"]

    master = {}
    for row in ws1.iter_rows(min_row=2, values_only=True):
        staff_id, name, site, rank = row[0], row[1], row[2], row[3]
        if not staff_id or not rank:
            continue
        sid = str(staff_id).strip()
        sid = B_TO_D.get(sid, sid)
        master[sid] = {
            "name": name, "site": site, "rank": rank,
            "hourly_wage": 0, "mgmt_fee": 0,
            "work_pattern": 5, "monthly_salary": None
        }

    for row in ws2.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        sid = str(row[0]).strip()
        sid = B_TO_D.get(sid, sid)
        wage = row[2]
        note = str(row[3]) if row[3] else ""
        if sid not in master:
            continue
        if "月給" in note:
            master[sid]["monthly_salary"] = int(str(wage).replace(",", "").strip()) if wage else 0
        else:
            master[sid]["hourly_wage"] = int(str(wage).replace(",", "").strip()) if wage else 0
        master[sid]["mgmt_fee"] = 3030 if "管理料" in note else 0

    return master

def jwt_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        token = request.headers.get("Authorization", "").replace("Bearer ", "")
        if not token:
            return jsonify({"error": "認証が必要です"}), 401
        try:
            payload = jwt.decode(token, JWT_SECRET, algorithms=["HS256"])
            request.staff_id = payload.get("staff_id")
            request.role = payload.get("role")
        except jwt.ExpiredSignatureError:
            return jsonify({"error": "トークンが期限切れです"}), 401
        except jwt.InvalidTokenError:
            return jsonify({"error": "無効なトークンです"}), 401
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        token = request.headers.get("Authorization", "").replace("Bearer ", "")
        if not token:
            return jsonify({"error": "認証が必要です"}), 401
        try:
            payload = jwt.decode(token, JWT_SECRET, algorithms=["HS256"])
            if payload.get("role") != "admin":
                return jsonify({"error": "管理者権限が必要です"}), 403
            request.staff_id = payload.get("staff_id")
            request.role = payload.get("role")
        except jwt.InvalidTokenError:
            return jsonify({"error": "無効なトークンです"}), 401
        return f(*args, **kwargs)
    return decorated

def register_staff_routes(app):

    @app.route("/health_staff")
    def health_staff():
        return jsonify({"status": "ok", "service": "staff-dashboard"})

    @app.route("/staff/init_admin")
    def init_admin():
        try:
            password = "Dghojin2026"
            password_hash = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
            supabase_staff.table("staff_master").upsert({
                "staff_id": "ADMIN001",
                "staff_name": "管理者",
                "login_id": "admin",
                "password_hash": password_hash,
                "role": "admin"
            }).execute()
            return jsonify({"status": "ok", "message": "管理者アカウントを作成しました"})
        except Exception as e:
            import traceback
            return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500

    @app.route("/staff/login", methods=["POST"])
    def staff_login():
        try:
            data = request.get_json()
            login_id = data.get("login_id", "").strip()
            password = data.get("password", "").strip()
            if not login_id or not password:
                return jsonify({"error": "IDとパスワードを入力してください"}), 400
            res = supabase_staff.table("staff_master")\
                .select("*").eq("login_id", login_id).execute()
            if not res.data:
                return jsonify({"error": "IDまたはパスワードが間違っています"}), 401
            staff = res.data[0]
            if not bcrypt.checkpw(password.encode(), staff["password_hash"].encode()):
                return jsonify({"error": "IDまたはパスワードが間違っています"}), 401
            token = jwt.encode({
                "staff_id": staff["staff_id"],
                "role": staff["role"],
                "exp": datetime.datetime.utcnow() + datetime.timedelta(days=30)
            }, JWT_SECRET, algorithm="HS256")
            return jsonify({
                "status": "ok",
                "token": token,
                "staff_id": staff["staff_id"],
                "role": staff["role"],
                "name": staff["staff_name"]
            })
        except Exception as e:
            import traceback
            return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500

    @app.route("/staff/register", methods=["POST"])
    @admin_required
    def staff_register():
        try:
            data = request.get_json()
            staff_id = data.get("staff_id", "").strip()
            staff_name = data.get("staff_name", "").strip()
            login_id = data.get("login_id", "").strip()
            password = data.get("password", "").strip()
            role = data.get("role", "staff")
            if not all([staff_id, staff_name, login_id, password]):
                return jsonify({"error": "必須項目が不足しています"}), 400
            password_hash = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
            supabase_staff.table("staff_master").upsert({
                "staff_id": staff_id,
                "staff_name": staff_name,
                "login_id": login_id,
                "password_hash": password_hash,
                "role": role
            }).execute()
            return jsonify({"status": "ok", "staff_id": staff_id})
        except Exception as e:
            import traceback
            return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500

    @app.route("/staff/me")
    @jwt_required
    def staff_me():
        try:
            res = supabase_staff.table("staff_master")\
                .select("staff_id,staff_name,role")\
                .eq("staff_id", request.staff_id).execute()
            if not res.data:
                return js
