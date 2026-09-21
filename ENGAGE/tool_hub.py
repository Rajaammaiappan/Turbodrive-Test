"""
tool_hub.py
=========================================================================
ENGAGE - Engineering Gateway for Automation, Guidance & Enterprise Tools
(RR / ALTEN internal tool)   Tagline: "Launch. Automate. Accelerate."

Single entry point / dashboard for all internally developed engineering
tools. Lets users discover, launch, learn about, and monitor every tool
from one place. Tracks usage so management can see which tools are used,
how often, and estimated hours saved.

Runs on Flask 1.1.2 / Python 3.9 (corporate Anaconda) - port 5010.

Run:
    "C:\\ProgramData\\Anaconda3\\python.exe" tool_hub.py
or double-click run_tool_hub.bat, then the browser opens automatically.

Data files (created next to this script if missing):
    tools_config.json   -> the list of tools (editable via Admin page)
    hub_data.db          -> sqlite: usage_events, favorites
=========================================================================
"""

import base64
import csv
import io
import json
import mimetypes
import os
import platform
import sqlite3
import subprocess
import threading
import time
import traceback
import webbrowser
from datetime import datetime, timedelta
from pathlib import Path

from flask import (Flask, jsonify, render_template_string, request,
                    send_file)
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

app = Flask(__name__)
PORT = 5010

BASE_DIR = Path(__file__).parent.resolve()
CONFIG_PATH = BASE_DIR / "tools_config.json"
HUB_DB_PATH = BASE_DIR / "hub_data.db"
LOGOS_DIR = BASE_DIR / "logos"

# ------------------------------------------------------------------ #
#  Logos - embedded as base64 data URIs so the header always renders,
#  even with no internet access (corporate machines are often offline).
#  Drop the real files in a "logos" folder next to this script:
#     logos/rr_logo.png  (or .svg)
#     logos/alten_logo.png
#  If a file isn't found, the header falls back to the RR/ALTEN text
#  badges so the app never breaks.
# ------------------------------------------------------------------ #
_logo_cache = {}


def logo_data_uri(*candidate_names):
    """Return a base64 data: URI for the first matching file found in
    logos/, or None if none of the candidates exist."""
    cache_key = candidate_names
    if cache_key in _logo_cache:
        return _logo_cache[cache_key]
    for name in candidate_names:
        path = LOGOS_DIR / name
        if path.exists() and path.is_file():
            mime, _ = mimetypes.guess_type(str(path))
            mime = mime or "application/octet-stream"
            data = base64.b64encode(path.read_bytes()).decode("ascii")
            uri = f"data:{mime};base64,{data}"
            _logo_cache[cache_key] = uri
            return uri
    _logo_cache[cache_key] = None
    return None


def get_logos():
    return {
        "rr_logo": logo_data_uri("rr_logo.svg", "rr_logo.png", "rr-logo.svg", "rr-logo.png"),
        "alten_logo": logo_data_uri("alten_logo.png", "alten_logo.svg",
                                    "alten-logo.png", "favicon-alten.png"),
    }

# Simple admin passcode (change this). Not real security - just keeps
# casual users out of the admin screen on a shared machine.
ADMIN_PASSCODE = "admin123"

# ------------------------------------------------------------------ #
#  Config (tools) storage - JSON file, no code changes needed to add
#  a new tool.
# ------------------------------------------------------------------ #
_config_lock = threading.Lock()


def load_tools():
    if not CONFIG_PATH.exists():
        return []
    with _config_lock:
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return []


def save_tools(tools):
    with _config_lock:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(tools, f, indent=2, ensure_ascii=False)


def next_tool_id(tools, tool_name):
    base = "".join(ch.lower() if ch.isalnum() else "_" for ch in tool_name).strip("_") or "tool"
    existing = {t["id"] for t in tools}
    candidate = base
    n = 1
    while candidate in existing:
        n += 1
        candidate = f"{base}_{n}"
    return candidate


# ------------------------------------------------------------------ #
#  Opening tools / guides - two kinds of "location" are supported:
#    1. A real web link (http:// or https://)      -> opened via the
#       browser (window.open) - unchanged behaviour.
#    2. A local path or UNC network path (.bat, .exe, .pdf, .docx,
#       .html, a SharePoint-synced folder path, etc.) -> a browser
#       CANNOT launch these directly (window.open only works for real
#       URLs, never for backslash paths or local executables - this is
#       a browser security restriction, not a bug in the hub). Since
#       this hub runs locally on the user's own machine (127.0.0.1),
#       the Flask backend opens these the same way Explorer would when
#       you double-click them: with the OS's default handler. That
#       means .bat/.exe get launched, and .pdf/.docx/.html open in
#       whatever app is set as default for that file type.
# ------------------------------------------------------------------ #
def looks_like_web_url(value):
    v = (value or "").strip().lower()
    return v.startswith("http://") or v.startswith("https://")


def open_local_path(path):
    """Open a local/UNC path with the OS default handler, exactly like
    double-clicking it in Explorer. Returns (ok, error_message)."""
    path = (path or "").strip()
    if not path:
        return False, "No location configured for this tool."
    try:
        if platform.system() == "Windows":
            os.startfile(path)  # type: ignore[attr-defined]  (Windows-only)
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
        return True, None
    except FileNotFoundError:
        return False, f"Path not found or not reachable from this PC: {path}"
    except OSError as exc:
        return False, f"Could not open this path ({exc})."


TOOL_DEFAULTS = {
    "category": "Engineering", "icon": "\U0001F527", "description": "",
    "owner": "", "version": "1.0",
    "last_updated": datetime.now().strftime("%Y-%m-%d"),
    "status": "active", "tool_url": "", "guide_url": "",
    "db_type": "none", "db_path": "", "db_viewer_url": "",
    "hours_saved_per_use": 0,
    "purpose": "", "features": [], "benefits": [], "scope": "",
    "contact": "", "change_history": [],
}


# ------------------------------------------------------------------ #
#  Hub DB - usage tracking + favorites (this hub's own sqlite db,
#  separate from each tool's own database)
# ------------------------------------------------------------------ #
def get_hub_con():
    con = sqlite3.connect(str(HUB_DB_PATH), timeout=15)
    con.row_factory = sqlite3.Row
    return con


def init_hub_db():
    con = get_hub_con()
    con.execute("""CREATE TABLE IF NOT EXISTS usage_events (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tool_id TEXT NOT NULL,
        action TEXT NOT NULL,
        user TEXT,
        ts TEXT NOT NULL
    )""")
    con.execute("""CREATE TABLE IF NOT EXISTS favorites (
        user TEXT NOT NULL,
        tool_id TEXT NOT NULL,
        PRIMARY KEY (user, tool_id)
    )""")
    con.execute("""CREATE TABLE IF NOT EXISTS tool_audit (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tool_id TEXT NOT NULL,
        tool_name TEXT,
        change_type TEXT NOT NULL,
        from_version TEXT,
        to_version TEXT,
        field_changes TEXT,
        notes TEXT,
        changed_by TEXT,
        ts TEXT NOT NULL,
        snapshot TEXT
    )""")
    con.commit()
    con.close()


def record_audit(tool_id, tool_name, change_type, from_version, to_version,
                  field_changes, notes, changed_by, snapshot=None):
    con = get_hub_con()
    con.execute(
        "INSERT INTO tool_audit (tool_id, tool_name, change_type, from_version, "
        "to_version, field_changes, notes, changed_by, ts, snapshot) "
        "VALUES (?,?,?,?,?,?,?,?,?,?)",
        (tool_id, tool_name, change_type, from_version, to_version,
         json.dumps(field_changes, ensure_ascii=False),
         (notes or "").strip(), (changed_by or "admin").strip() or "admin",
         datetime.now().isoformat(timespec="seconds"),
         json.dumps(snapshot, ensure_ascii=False) if snapshot is not None else None),
    )
    con.commit()
    con.close()


def diff_tool_fields(old, new):
    """Return list of {field, old, new} for every TOOL_DEFAULTS-tracked field
    (plus tool_name) that changed between old and new tool dicts."""
    changes = []
    tracked = list(TOOL_DEFAULTS.keys()) + ["tool_name"]
    for field in tracked:
        old_val = old.get(field)
        new_val = new.get(field)
        if old_val != new_val:
            changes.append({"field": field, "old": old_val, "new": new_val})
    return changes


def log_event(tool_id, action, user):
    con = get_hub_con()
    con.execute(
        "INSERT INTO usage_events (tool_id, action, user, ts) VALUES (?,?,?,?)",
        (tool_id, action, (user or "anonymous").strip() or "anonymous",
         datetime.now().isoformat(timespec="seconds")),
    )
    con.commit()
    con.close()


# ------------------------------------------------------------------ #
#  Stats / analytics helpers
# ------------------------------------------------------------------ #
def get_tool_db_usage(tool):
    """Read the REAL usage count (and, if possible, distinct users)
    straight from the tool's own database — the authoritative record of
    how many times the tool actually ran, not just how many times
    someone clicked 'Open' inside the hub. Returns (count, users) where
    both are None if the tool has no reachable/configured db, so
    callers can fall back to the hub's own click log for that tool."""
    path = resolve_db_path(tool)
    if not path or not path.exists():
        return None, None
    try:
        con = sqlite3.connect(f"file:{path}?mode=ro", uri=True, timeout=5)
    except sqlite3.Error:
        return None, None
    table = pick_best_table(con)
    if not table:
        con.close()
        return None, None
    try:
        total = con.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0]
    except sqlite3.Error:
        con.close()
        return None, None
    cols = [r[1] for r in con.execute(f'PRAGMA table_info("{table}")').fetchall()]
    user_col = next((c for c in cols if c.lower() in
                      ("downloaded_by", "user", "username", "created_by",
                       "requested_by", "owner", "processed_by")), None)
    users = set()
    if user_col:
        try:
            users = {r[0] for r in con.execute(
                f'SELECT DISTINCT "{user_col}" FROM "{table}" '
                f'WHERE "{user_col}" IS NOT NULL AND "{user_col}" != \'\'').fetchall()}
        except sqlite3.Error:
            users = set()
    con.close()
    return total, users


def build_stats_and_analytics():
    tools = load_tools()
    tool_map = {t["id"]: t for t in tools}
    con = get_hub_con()
    rows = con.execute("SELECT tool_id, action, user, ts FROM usage_events").fetchall()
    con.close()

    total_tools = len(tools)
    active_tools = len([t for t in tools if t.get("status", "active") == "active"])
    open_events = [r for r in rows if r["action"] == "open"]
    hub_click_counts = {}
    for r in open_events:
        hub_click_counts[r["tool_id"]] = hub_click_counts.get(r["tool_id"], 0) + 1

    distinct_users = {r["user"] for r in rows if r["user"]}

    # Real usage count per tool: prefer the actual record count from the
    # tool's own database (its true number of runs) over the hub's click
    # log, since not every "Open" click necessarily ran the tool and not
    # every run of the tool necessarily went through the hub. Hours saved
    # is then simply (real usage count x hours saved per use), summed.
    per_tool_counts = {}
    hours_saved = 0.0
    for t in tools:
        db_count, db_users = get_tool_db_usage(t)
        if db_count is not None:
            count = db_count
            if db_users:
                distinct_users |= db_users
        else:
            count = hub_click_counts.get(t["id"], 0)
        per_tool_counts[t["id"]] = count
        hours_saved += count * float(t.get("hours_saved_per_use", 0) or 0)

    total_automations = sum(per_tool_counts.values())

    most_used = sorted(per_tool_counts.items(), key=lambda kv: kv[1], reverse=True)[:8]
    most_used_named = [
        {"tool_id": tid, "tool_name": tool_map.get(tid, {}).get("tool_name", tid), "count": c}
        for tid, c in most_used if c > 0
    ]

    recent = sorted(rows, key=lambda r: r["ts"], reverse=True)[:25]
    recent_named = [
        {
            "tool_id": r["tool_id"],
            "tool_name": tool_map.get(r["tool_id"], {}).get("tool_name", r["tool_id"]),
            "action": r["action"], "user": r["user"], "ts": r["ts"],
        }
        for r in recent
    ]

    # category breakdown for the analytics chart
    cat_counts = {}
    for t in tools:
        cat_counts[t.get("category", "Other")] = cat_counts.get(t.get("category", "Other"), 0) + 1

    # usage over last 14 days (hub click activity only - per-tool DBs don't
    # all expose a reliable per-day timestamp column across every tool)
    since = datetime.now() - timedelta(days=14)
    daily = {}
    for r in rows:
        try:
            d = datetime.fromisoformat(r["ts"])
        except ValueError:
            continue
        if d < since:
            continue
        key = d.strftime("%Y-%m-%d")
        daily[key] = daily.get(key, 0) + 1
    daily_series = []
    for i in range(13, -1, -1):
        day = (datetime.now() - timedelta(days=i)).strftime("%Y-%m-%d")
        daily_series.append({"date": day, "count": daily.get(day, 0)})

    stats = {
        "total_tools": total_tools,
        "active_tools": active_tools,
        "active_users": len(distinct_users),
        "total_automations": total_automations,
        "hours_saved": round(hours_saved, 1),
    }
    analytics = {
        "most_used": most_used_named,
        "recent": recent_named,
        "categories": [{"category": k, "count": v} for k, v in cat_counts.items()],
        "daily": daily_series,
        "per_tool_counts": per_tool_counts,
    }
    return stats, analytics


# ------------------------------------------------------------------ #
#  Per-tool database viewer (generic - works on any sqlite db)
# ------------------------------------------------------------------ #
def resolve_db_path(tool):
    p = (tool.get("db_path") or "").strip()
    if not p:
        return None
    path = Path(p)
    if not path.is_absolute():
        path = BASE_DIR / path
    return path


def pick_best_table(con):
    names = [r[0] for r in con.execute(
        "SELECT name FROM sqlite_master WHERE type='table' "
        "AND name NOT LIKE 'sqlite_%'").fetchall()]
    if not names:
        return ""
    best, best_n = names[0], -1
    for n in names:
        try:
            c = con.execute(f'SELECT COUNT(*) FROM "{n}"').fetchone()[0]
        except sqlite3.Error:
            c = 0
        if c > best_n:
            best, best_n = n, c
    return best


def tool_db_query(tool, q, page, page_size=50):
    """Return dict with columns, rows, total, page, table, recent_activity."""
    path = resolve_db_path(tool)
    if not path or not path.exists():
        return {"ok": False, "error": "Database not found for this tool.",
                "columns": [], "rows": [], "total": 0, "table": ""}
    try:
        con = sqlite3.connect(f"file:{path}?mode=ro", uri=True, timeout=10)
        con.row_factory = sqlite3.Row
    except sqlite3.Error as exc:
        return {"ok": False, "error": f"Could not open database ({exc})",
                "columns": [], "rows": [], "total": 0, "table": ""}

    table = pick_best_table(con)
    if not table:
        con.close()
        return {"ok": False, "error": "No tables found in this database.",
                "columns": [], "rows": [], "total": 0, "table": ""}

    cols = [r[1] for r in con.execute(f'PRAGMA table_info("{table}")').fetchall()]
    where, params = [], []
    q = (q or "").strip()
    if q and cols:
        where.append("(" + " OR ".join(f'CAST("{c}" AS TEXT) LIKE ?' for c in cols) + ")")
        params += [f"%{q}%"] * len(cols)

    where_sql = (" WHERE " + " AND ".join(where)) if where else ""
    total = con.execute(f'SELECT COUNT(*) FROM "{table}"{where_sql}', params).fetchone()[0]

    date_col = next((c for c in cols if c.lower() in
                      ("processed_at", "created_at", "timestamp", "date", "ts")), None)
    order_sql = f' ORDER BY "{date_col}" DESC' if date_col else ""

    offset = max(page, 0) * page_size
    sql = f'SELECT * FROM "{table}"{where_sql}{order_sql} LIMIT ? OFFSET ?'
    rows = [dict(r) for r in con.execute(sql, params + [page_size, offset]).fetchall()]

    recent_activity = []
    if date_col:
        recent_activity = [r[date_col] for r in
                            con.execute(f'SELECT "{date_col}" FROM "{table}" '
                                        f'ORDER BY "{date_col}" DESC LIMIT 5').fetchall()]

    con.close()
    return {"ok": True, "columns": cols, "rows": rows, "total": total,
            "table": table, "recent_activity": recent_activity,
            "page": page, "page_size": page_size}


def tool_db_all_matching(tool, q):
    """Fetch ALL matching rows (no pagination) for export."""
    path = resolve_db_path(tool)
    if not path or not path.exists():
        return [], []
    con = sqlite3.connect(f"file:{path}?mode=ro", uri=True, timeout=10)
    con.row_factory = sqlite3.Row
    table = pick_best_table(con)
    if not table:
        con.close()
        return [], []
    cols = [r[1] for r in con.execute(f'PRAGMA table_info("{table}")').fetchall()]
    where, params = [], []
    q = (q or "").strip()
    if q and cols:
        where.append("(" + " OR ".join(f'CAST("{c}" AS TEXT) LIKE ?' for c in cols) + ")")
        params += [f"%{q}%"] * len(cols)
    where_sql = (" WHERE " + " AND ".join(where)) if where else ""
    rows = [dict(r) for r in con.execute(
        f'SELECT * FROM "{table}"{where_sql}', params).fetchall()]
    con.close()
    return cols, rows


# ------------------------------------------------------------------ #
#  Routes - pages
# ------------------------------------------------------------------ #
@app.route("/")
def index():
    return render_template_string(PAGE_DASHBOARD, **get_logos())


@app.route("/admin")
def admin_page():
    return render_template_string(PAGE_ADMIN, **get_logos())


# ------------------------------------------------------------------ #
#  Routes - API: tools / stats / analytics / favorites / logging
# ------------------------------------------------------------------ #
@app.route("/api/tools")
def api_tools():
    tools = load_tools()
    _, analytics = build_stats_and_analytics()
    counts = analytics["per_tool_counts"]
    for t in tools:
        t["usage_count"] = counts.get(t["id"], 0)
    return jsonify(tools=tools)


@app.route("/api/stats")
def api_stats():
    stats, _ = build_stats_and_analytics()
    return jsonify(stats)


@app.route("/api/analytics")
def api_analytics():
    _, analytics = build_stats_and_analytics()
    return jsonify(analytics)


@app.route("/api/log", methods=["POST"])
def api_log():
    data = request.get_json(force=True, silent=True) or {}
    tool_id = data.get("tool_id")
    action = data.get("action")
    user = data.get("user")
    if not tool_id or action not in ("open", "guide", "about", "dashboard"):
        return jsonify(ok=False, error="Invalid log request"), 400
    log_event(tool_id, action, user)
    return jsonify(ok=True)


@app.route("/api/resolve_open", methods=["POST"])
def api_resolve_open():
    """Used by the 'Open Tool' and 'User Guide' buttons. If the tool's
    configured location is a real web link, hands it back so the
    frontend opens it in a new tab. If it's a local/UNC path, opens it
    here on the backend with the OS default handler (see open_local_path)."""
    data = request.get_json(force=True, silent=True) or {}
    tool_id = data.get("tool_id")
    field = data.get("field")
    if field not in ("tool_url", "guide_url"):
        return jsonify(ok=False, error="Invalid field"), 400
    tools = load_tools()
    tool = next((t for t in tools if t["id"] == tool_id), None)
    if not tool:
        return jsonify(ok=False, error="Unknown tool"), 404
    target = (tool.get(field) or "").strip()
    label = "Tool URL" if field == "tool_url" else "User Guide"
    if not target:
        return jsonify(ok=False, error=f"No {label} configured for this tool yet.")
    if looks_like_web_url(target):
        return jsonify(ok=True, mode="url", url=target)
    ok, err = open_local_path(target)
    if ok:
        return jsonify(ok=True, mode="local")
    return jsonify(ok=False, error=err)


@app.route("/api/favorites")
def api_favorites_get():
    user = (request.args.get("user") or "anonymous").strip() or "anonymous"
    con = get_hub_con()
    rows = con.execute("SELECT tool_id FROM favorites WHERE user=?", (user,)).fetchall()
    con.close()
    return jsonify(favorites=[r["tool_id"] for r in rows])


@app.route("/api/favorites/toggle", methods=["POST"])
def api_favorites_toggle():
    data = request.get_json(force=True, silent=True) or {}
    user = (data.get("user") or "anonymous").strip() or "anonymous"
    tool_id = data.get("tool_id")
    if not tool_id:
        return jsonify(ok=False, error="tool_id required"), 400
    con = get_hub_con()
    existing = con.execute("SELECT 1 FROM favorites WHERE user=? AND tool_id=?",
                            (user, tool_id)).fetchone()
    if existing:
        con.execute("DELETE FROM favorites WHERE user=? AND tool_id=?", (user, tool_id))
        is_fav = False
    else:
        con.execute("INSERT INTO favorites (user, tool_id) VALUES (?,?)", (user, tool_id))
        is_fav = True
    con.commit()
    con.close()
    return jsonify(ok=True, favorite=is_fav)


# ------------------------------------------------------------------ #
#  Routes - API: per-tool data dashboard / db viewer
# ------------------------------------------------------------------ #
@app.route("/api/tool_db/<tool_id>")
def api_tool_db(tool_id):
    tools = load_tools()
    tool = next((t for t in tools if t["id"] == tool_id), None)
    if not tool:
        return jsonify(ok=False, error="Unknown tool"), 404
    q = request.args.get("q", "")
    page = int(request.args.get("page", 0) or 0)
    result = tool_db_query(tool, q, page)
    return jsonify(result)


@app.route("/api/tool_db/<tool_id>/export")
def api_tool_db_export(tool_id):
    fmt = (request.args.get("format") or "xlsx").lower()
    q = request.args.get("q", "")
    tools = load_tools()
    tool = next((t for t in tools if t["id"] == tool_id), None)
    if not tool:
        return jsonify(ok=False, error="Unknown tool"), 404
    cols, rows = tool_db_all_matching(tool, q)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    fname = f"{tool_id}_export_{stamp}.{fmt if fmt in ('xlsx', 'csv') else 'xlsx'}"

    if fmt == "csv":
        buf = io.StringIO()
        writer = csv.writer(buf)
        writer.writerow(cols)
        for r in rows:
            writer.writerow([r.get(c, "") for c in cols])
        mem = io.BytesIO(buf.getvalue().encode("utf-8-sig"))
        return send_file(mem, as_attachment=True, download_name=fname,
                          mimetype="text/csv")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(cols)
    header_fill = PatternFill(start_color="1A4FAD", end_color="1A4FAD", fill_type="solid")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
    for r in rows:
        ws.append([r.get(c, "") for c in cols])
    for i, col in enumerate(cols, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 18
    mem = io.BytesIO()
    wb.save(mem)
    mem.seek(0)
    return send_file(mem, as_attachment=True, download_name=fname,
                      mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ------------------------------------------------------------------ #
#  Routes - API: admin (add / edit / delete / enable-disable / usage)
# ------------------------------------------------------------------ #
def check_admin(data_or_args):
    return (data_or_args.get("passcode") or "") == ADMIN_PASSCODE


@app.route("/api/admin/login", methods=["POST"])
def api_admin_login():
    data = request.get_json(force=True, silent=True) or {}
    return jsonify(ok=check_admin(data))


@app.route("/api/admin/tools", methods=["POST"])
def api_admin_add_tool():
    data = request.get_json(force=True, silent=True) or {}
    if not check_admin(data):
        return jsonify(ok=False, error="Invalid passcode"), 403
    tool_name = (data.get("tool_name") or "").strip()
    if not tool_name:
        return jsonify(ok=False, error="tool_name is required"), 400
    tools = load_tools()
    new_tool = dict(TOOL_DEFAULTS)
    new_tool.update({k: v for k, v in data.items() if k in TOOL_DEFAULTS or k == "tool_name"})
    new_tool["tool_name"] = tool_name
    new_tool["id"] = next_tool_id(tools, tool_name)
    new_tool["change_history"] = [{
        "version": new_tool.get("version", "1.0"),
        "date": datetime.now().strftime("%Y-%m-%d"),
        "notes": (data.get("change_notes") or "Initial version added to hub.").strip(),
    }]
    tools.append(new_tool)
    save_tools(tools)
    record_audit(new_tool["id"], tool_name, "created", None, new_tool.get("version"),
                 [], data.get("change_notes"), data.get("changed_by"), snapshot=new_tool)
    return jsonify(ok=True, tool=new_tool)


@app.route("/api/admin/tools/<tool_id>", methods=["PUT"])
def api_admin_edit_tool(tool_id):
    data = request.get_json(force=True, silent=True) or {}
    if not check_admin(data):
        return jsonify(ok=False, error="Invalid passcode"), 403
    tools = load_tools()
    idx = next((i for i, t in enumerate(tools) if t["id"] == tool_id), None)
    if idx is None:
        return jsonify(ok=False, error="Tool not found"), 404

    old_tool = json.loads(json.dumps(tools[idx]))  # deep copy for diff + rollback snapshot
    old_version = old_tool.get("version")

    for k, v in data.items():
        if k in TOOL_DEFAULTS or k == "tool_name":
            tools[idx][k] = v
    tools[idx]["last_updated"] = datetime.now().strftime("%Y-%m-%d")
    new_version = tools[idx].get("version")

    changes = diff_tool_fields(old_tool, tools[idx])
    change_notes = (data.get("change_notes") or "").strip()

    # Version control: whenever the version number changes (or notes were
    # supplied even without a version bump), record it in the tool's own
    # change_history so it shows in the About popup, and in the hub-wide
    # audit trail for admins.
    if changes:
        history = tools[idx].get("change_history") or []
        if new_version != old_version or change_notes:
            entry_notes = change_notes or f"Updated ({len(changes)} field(s) changed)."
            history.insert(0, {
                "version": new_version or old_version or "1.0",
                "date": datetime.now().strftime("%Y-%m-%d"),
                "notes": entry_notes,
            })
            tools[idx]["change_history"] = history

    save_tools(tools)

    if changes:
        record_audit(tool_id, tools[idx].get("tool_name"), "edit", old_version, new_version,
                     changes, change_notes, data.get("changed_by"), snapshot=old_tool)
    return jsonify(ok=True, tool=tools[idx], changed_fields=len(changes))


@app.route("/api/admin/tools/<tool_id>", methods=["DELETE"])
def api_admin_delete_tool(tool_id):
    passcode = request.args.get("passcode", "")
    changed_by = request.args.get("changed_by", "admin")
    if not check_admin({"passcode": passcode}):
        return jsonify(ok=False, error="Invalid passcode"), 403
    tools = load_tools()
    removed = next((t for t in tools if t["id"] == tool_id), None)
    tools = [t for t in tools if t["id"] != tool_id]
    save_tools(tools)
    if removed:
        record_audit(tool_id, removed.get("tool_name"), "deleted",
                     removed.get("version"), None, [], "Tool removed from hub.",
                     changed_by, snapshot=removed)
    return jsonify(ok=True)


@app.route("/api/admin/tools/<tool_id>/toggle", methods=["POST"])
def api_admin_toggle_tool(tool_id):
    data = request.get_json(force=True, silent=True) or {}
    if not check_admin(data):
        return jsonify(ok=False, error="Invalid passcode"), 403
    tools = load_tools()
    idx = next((i for i, t in enumerate(tools) if t["id"] == tool_id), None)
    if idx is None:
        return jsonify(ok=False, error="Tool not found"), 404
    old_status = tools[idx].get("status")
    tools[idx]["status"] = "disabled" if old_status == "active" else "active"
    save_tools(tools)
    record_audit(tool_id, tools[idx].get("tool_name"), "status_toggle", None, None,
                 [{"field": "status", "old": old_status, "new": tools[idx]["status"]}],
                 data.get("change_notes"), data.get("changed_by"))
    return jsonify(ok=True, status=tools[idx]["status"])


@app.route("/api/admin/tools/<tool_id>/history")
def api_admin_tool_history(tool_id):
    passcode = request.args.get("passcode", "")
    if not check_admin({"passcode": passcode}):
        return jsonify(ok=False, error="Invalid passcode"), 403
    con = get_hub_con()
    rows = con.execute(
        "SELECT * FROM tool_audit WHERE tool_id=? ORDER BY ts DESC", (tool_id,)
    ).fetchall()
    con.close()
    audit = []
    for r in rows:
        audit.append({
            "id": r["id"], "change_type": r["change_type"],
            "from_version": r["from_version"], "to_version": r["to_version"],
            "field_changes": json.loads(r["field_changes"] or "[]"),
            "notes": r["notes"], "changed_by": r["changed_by"], "ts": r["ts"],
            "has_snapshot": r["snapshot"] is not None,
        })
    tools = load_tools()
    tool = next((t for t in tools if t["id"] == tool_id), None)
    change_history = (tool or {}).get("change_history", [])
    return jsonify(ok=True, audit=audit, change_history=change_history)


@app.route("/api/admin/tools/<tool_id>/rollback", methods=["POST"])
def api_admin_tool_rollback(tool_id):
    data = request.get_json(force=True, silent=True) or {}
    if not check_admin(data):
        return jsonify(ok=False, error="Invalid passcode"), 403
    audit_id = data.get("audit_id")
    if not audit_id:
        return jsonify(ok=False, error="audit_id is required"), 400
    con = get_hub_con()
    row = con.execute("SELECT * FROM tool_audit WHERE id=? AND tool_id=?",
                       (audit_id, tool_id)).fetchone()
    con.close()
    if not row or not row["snapshot"]:
        return jsonify(ok=False, error="No snapshot available to roll back to."), 404
    snapshot = json.loads(row["snapshot"])

    tools = load_tools()
    idx = next((i for i, t in enumerate(tools) if t["id"] == tool_id), None)
    if idx is None:
        return jsonify(ok=False, error="Tool not found"), 404

    current = json.loads(json.dumps(tools[idx]))
    restored = dict(current)
    for k in TOOL_DEFAULTS:
        if k in snapshot:
            restored[k] = snapshot[k]
    restored["tool_name"] = snapshot.get("tool_name", current.get("tool_name"))
    restored["last_updated"] = datetime.now().strftime("%Y-%m-%d")
    history = restored.get("change_history") or []
    history.insert(0, {
        "version": restored.get("version"),
        "date": datetime.now().strftime("%Y-%m-%d"),
        "notes": f"Rolled back to a previous version (v{snapshot.get('version')}).",
    })
    restored["change_history"] = history
    tools[idx] = restored
    save_tools(tools)

    changes = diff_tool_fields(current, restored)
    record_audit(tool_id, restored.get("tool_name"), "rollback",
                 current.get("version"), restored.get("version"), changes,
                 f"Rolled back to version from audit #{audit_id}.",
                 data.get("changed_by"), snapshot=current)
    return jsonify(ok=True, tool=restored)


@app.route("/api/admin/usage")
def api_admin_usage():
    passcode = request.args.get("passcode", "")
    if not check_admin({"passcode": passcode}):
        return jsonify(ok=False, error="Invalid passcode"), 403
    tools = load_tools()
    tool_map = {t["id"]: t for t in tools}
    con = get_hub_con()
    rows = con.execute(
        "SELECT tool_id, action, COUNT(*) as n FROM usage_events GROUP BY tool_id, action"
    ).fetchall()
    con.close()
    per_tool = {}
    for r in rows:
        d = per_tool.setdefault(r["tool_id"], {
            "tool_id": r["tool_id"],
            "tool_name": tool_map.get(r["tool_id"], {}).get("tool_name", r["tool_id"]),
            "open": 0, "guide": 0, "about": 0, "dashboard": 0})
        d[r["action"]] = r["n"]
    return jsonify(usage=list(per_tool.values()))


# ------------------------------------------------------------------ #
#  HTML templates are defined in tool_hub_templates.py and imported
#  below to keep this file manageable.
# ------------------------------------------------------------------ #
from tool_hub_templates import PAGE_DASHBOARD, PAGE_ADMIN  # noqa: E402


if __name__ == "__main__":
    init_hub_db()
    if not CONFIG_PATH.exists():
        save_tools([])
    threading.Timer(1.5, lambda: webbrowser.open(f"http://127.0.0.1:{PORT}")).start()
    print(f"ENGAGE - Engineering Gateway for Automation, Guidance & Enterprise Tools")
    print(f"        \"Launch. Automate. Accelerate.\"")
    print(f"        http://127.0.0.1:{PORT}")
    print(f"Admin page:           http://127.0.0.1:{PORT}/admin")
    print(f"Config file:          {CONFIG_PATH}")
    app.run(host="127.0.0.1", port=PORT, debug=False, threaded=True, use_reloader=False)
