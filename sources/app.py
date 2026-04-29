#!/usr/bin/env python3
"""
SCM Vidi — Backend Flask
API REST pour partager les données entre utilisateurs YunoHost.
"""

import os
import io
import json
import base64
import hmac
import hashlib
import sqlite3
import time
from functools import wraps
from flask import Flask, request, jsonify, send_from_directory, g, Response

app = Flask(__name__, static_folder="static")

# ─── Config ───────────────────────────────────────────────────────────────────
DB_PATH    = os.environ.get("SCM_DB_PATH",    "/var/lib/scm_vidi/data.db")
ADMIN_USER = os.environ.get("SCM_ADMIN_USER", "admin")
SECRET_KEY = os.environ.get("SCM_SECRET_KEY", os.urandom(32).hex())

# ─── Base de données ──────────────────────────────────────────────────────────

def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
    return g.db

@app.teardown_appcontext
def close_db(e=None):
    db = g.pop("db", None)
    if db:
        db.close()

def init_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    db = sqlite3.connect(DB_PATH)
    db.execute("""
        CREATE TABLE IF NOT EXISTS app_state (
            key   TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
    """)
    db.execute("""
        CREATE TABLE IF NOT EXISTS history (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            username  TEXT NOT NULL,
            action    TEXT NOT NULL,
            ts        DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)
    db.commit()
    db.close()

# ─── Auth ─────────────────────────────────────────────────────────────────────

def make_token(username):
    expires = int(time.time()) + 86400
    payload = f"{username}:{expires}"
    sig = hmac.new(SECRET_KEY.encode(), payload.encode(), hashlib.sha256).hexdigest()
    return base64.b64encode(f"{payload}:{sig}".encode()).decode()

def verify_token(token):
    try:
        raw = base64.b64decode(token.encode()).decode()
        username, expires_str, sig = raw.rsplit(":", 2)
        if time.time() > int(expires_str):
            return None
        payload = f"{username}:{expires_str}"
        expected = hmac.new(SECRET_KEY.encode(), payload.encode(), hashlib.sha256).hexdigest()
        if hmac.compare_digest(sig, expected):
            return username
    except Exception:
        pass
    return None

def get_current_user():
    # Méthode 1 — Token signé (appels API JS)
    token = request.headers.get("X-SCM-Token", "")
    if token:
        user = verify_token(token)
        if user:
            return user
    # Méthode 2 — Authorization Basic (YunoHost 12)
    auth = request.headers.get("Authorization", "")
    if auth.startswith("Basic "):
        try:
            decoded = base64.b64decode(auth[6:]).decode("utf-8")
            username = decoded.split(":")[0]
            if username:
                return username
        except Exception:
            pass
    # Méthode 3 — X-Remote-User (YunoHost < 12)
    remote_user = request.headers.get("X-Remote-User", "")
    if remote_user:
        return remote_user
    return "anonymous"

def is_admin():
    return get_current_user() == ADMIN_USER

def require_admin(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not is_admin():
            return jsonify({"error": "Réservé à l'administrateur"}), 403
        return f(*args, **kwargs)
    return decorated

# ─── Helpers état ─────────────────────────────────────────────────────────────

def get_full_state():
    """Retourne l'état complet depuis la DB sous forme de dict."""
    db = get_db()
    rows = db.execute("SELECT key, value FROM app_state").fetchall()
    state = {}
    for row in rows:
        state[row["key"]] = json.loads(row["value"])
    # Reconstituer un objet état plat
    if state and "associes" in state:
        return state
    # Compatibilité : état stocké sous une clé "data"
    if "data" in state:
        return state["data"]
    return state

# ─── Routes API ───────────────────────────────────────────────────────────────

@app.route("/api/state", methods=["GET"])
def get_state():
    db = get_db()
    rows = db.execute("SELECT key, value FROM app_state").fetchall()
    state = {row["key"]: json.loads(row["value"]) for row in rows}
    return jsonify({
        "state": state,
        "user": get_current_user(),
        "is_admin": is_admin()
    })

@app.route("/api/state", methods=["POST"])
@require_admin
def save_state():
    data = request.get_json(force=True)
    if not isinstance(data, dict):
        return jsonify({"error": "JSON invalide"}), 400
    db = get_db()
    for key, value in data.items():
        db.execute(
            "INSERT OR REPLACE INTO app_state (key, value) VALUES (?, ?)",
            (key, json.dumps(value))
        )
    db.execute(
        "INSERT INTO history (username, action) VALUES (?, ?)",
        (get_current_user(), f"save_state keys={list(data.keys())}")
    )
    db.commit()
    return jsonify({"ok": True})

@app.route("/api/state/<key>", methods=["PUT"])
@require_admin
def update_key(key):
    value = request.get_json(force=True)
    db = get_db()
    db.execute(
        "INSERT OR REPLACE INTO app_state (key, value) VALUES (?, ?)",
        (key, json.dumps(value))
    )
    db.execute(
        "INSERT INTO history (username, action) VALUES (?, ?)",
        (get_current_user(), f"update_key key={key}")
    )
    db.commit()
    return jsonify({"ok": True})

@app.route("/api/history", methods=["GET"])
def get_history():
    db = get_db()
    rows = db.execute(
        "SELECT username, action, ts FROM history ORDER BY id DESC LIMIT 50"
    ).fetchall()
    return jsonify([dict(r) for r in rows])

@app.route("/api/whoami", methods=["GET"])
def whoami():
    return jsonify({"user": get_current_user(), "is_admin": is_admin()})

# ─── Export Excel ─────────────────────────────────────────────────────────────

@app.route("/api/export/excel/<int:annee>")
def export_excel(annee):
    """Génère et télécharge le fichier Excel pour l'année demandée."""
    try:
        import sys
        src_dir = os.path.dirname(os.path.abspath(__file__))
        if src_dir not in sys.path:
            sys.path.insert(0, src_dir)
        import export_excel as ex

        data = get_full_state()
        if not data or "associes" not in data:
            return jsonify({"error": "Aucune donnée disponible"}), 404

        from openpyxl import Workbook
        wb = Workbook()
        ex.build_recap(wb, data, annee)
        for mo in range(12):
            ex.build_mois(wb, data, annee, mo)
        assocs = [a for a in data["associes"] if a.get("actif")]
        for a in assocs:
            ex.build_ortho(wb, data, annee, a)

        # Sauvegarder dans un buffer mémoire
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        return Response(
            buf.read(),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f"attachment; filename=SCM_Vidi_{annee}.xlsx"
            }
        )
    except ImportError as e:
        return jsonify({"error": f"Dépendance manquante : {e}"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ─── Servir l'app HTML ────────────────────────────────────────────────────────

@app.route("/", defaults={"path": ""})
@app.route("/<path:path>")
def serve(path):
    if path and os.path.exists(os.path.join(app.static_folder, path)):
        return send_from_directory(app.static_folder, path)

    user  = get_current_user()
    admin = user == ADMIN_USER
    token = make_token(user) if user != "anonymous" else ""

    with open(os.path.join(app.static_folder, "index.html"), "r", encoding="utf-8") as f:
        html = f.read()

    injection = f"""<script>
window._YNH_USER     = "{user}";
window._YNH_IS_ADMIN = {"true" if admin else "false"};
window._YNH_TOKEN    = "{token}";
</script>"""
    html = html.replace("</head>", injection + "\n</head>", 1)
    return Response(html, mimetype="text/html")

# ─── Démarrage ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    init_db()
    port = int(os.environ.get("SCM_PORT", 5000))
    app.run(host="127.0.0.1", port=port)
