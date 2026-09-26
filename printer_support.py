"""Multi-printer database schema and REST API for the Flask ERP application."""
from __future__ import annotations

import base64
import json
import sqlite3
import uuid
from datetime import datetime
from typing import Any, Callable

from flask import jsonify, request, session

NOW = lambda: datetime.now().strftime("%Y-%m-%d %H:%M:%S")

PRINTER_KINDS = {"thermal", "barcode", "invoice"}
CONNECTIONS = {"network", "usb", "windows", "browser"}
JOB_TYPES = {"receipt", "barcode_label", "invoice", "report", "custom"}
PAYLOAD_FORMATS = {"esc_pos", "zpl", "tspl", "pdf", "raw", "html"}

SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS printer_profiles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    kind TEXT NOT NULL CHECK(kind IN ('thermal','barcode','invoice')),
    paper_width_mm REAL,
    label_width_mm REAL,
    label_height_mm REAL,
    dpi INTEGER,
    payload_format TEXT NOT NULL CHECK(payload_format IN ('esc_pos','zpl','tspl','pdf','raw','html')),
    driver_options TEXT NOT NULL DEFAULT '{}',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS printers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    profile_id INTEGER NOT NULL REFERENCES printer_profiles(id),
    name TEXT NOT NULL,
    code TEXT NOT NULL UNIQUE,
    connection TEXT NOT NULL CHECK(connection IN ('network','usb','windows','browser')),
    address TEXT,
    port INTEGER,
    windows_printer_name TEXT,
    status TEXT NOT NULL DEFAULT 'inactive' CHECK(status IN ('active','inactive','error','maintenance')),
    is_default INTEGER NOT NULL DEFAULT 0 CHECK(is_default IN (0,1)),
    enabled INTEGER NOT NULL DEFAULT 1 CHECK(enabled IN (0,1)),
    branch_id INTEGER,
    workstation_id TEXT,
    capabilities TEXT NOT NULL DEFAULT '{}',
    last_seen_at TEXT,
    last_error TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    FOREIGN KEY(profile_id) REFERENCES printer_profiles(id)
);
CREATE UNIQUE INDEX IF NOT EXISTS one_default_printer_per_kind
ON printers(profile_id) WHERE is_default=1 AND enabled=1;
CREATE INDEX IF NOT EXISTS printers_lookup_idx
ON printers(enabled, status, branch_id, workstation_id);
CREATE TABLE IF NOT EXISTS printer_assignments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    printer_id INTEGER NOT NULL REFERENCES printers(id) ON DELETE CASCADE,
    job_type TEXT NOT NULL CHECK(job_type IN ('receipt','barcode_label','invoice','report','custom')),
    branch_id INTEGER,
    workstation_id TEXT,
    priority INTEGER NOT NULL DEFAULT 100,
    enabled INTEGER NOT NULL DEFAULT 1 CHECK(enabled IN (0,1)),
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    UNIQUE(printer_id, job_type, branch_id, workstation_id)
);
CREATE INDEX IF NOT EXISTS printer_assignment_lookup_idx
ON printer_assignments(job_type, branch_id, workstation_id, enabled, priority);
CREATE TABLE IF NOT EXISTS print_jobs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    printer_id INTEGER REFERENCES printers(id) ON DELETE SET NULL,
    job_type TEXT NOT NULL CHECK(job_type IN ('receipt','barcode_label','invoice','report','custom')),
    status TEXT NOT NULL DEFAULT 'queued' CHECK(status IN ('queued','printing','completed','failed','cancelled')),
    payload_format TEXT NOT NULL CHECK(payload_format IN ('esc_pos','zpl','tspl','pdf','raw','html')),
    payload TEXT,
    raw_payload BLOB,
    copies INTEGER NOT NULL DEFAULT 1 CHECK(copies BETWEEN 1 AND 100),
    idempotency_key TEXT UNIQUE,
    source_type TEXT,
    source_id TEXT,
    requested_by INTEGER,
    branch_id INTEGER,
    workstation_id TEXT,
    attempts INTEGER NOT NULL DEFAULT 0,
    error_message TEXT,
    queued_at TEXT NOT NULL,
    started_at TEXT,
    completed_at TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    CHECK(payload IS NOT NULL OR raw_payload IS NOT NULL)
);
CREATE INDEX IF NOT EXISTS print_jobs_queue_idx ON print_jobs(status, queued_at);
CREATE INDEX IF NOT EXISTS print_jobs_printer_idx ON print_jobs(printer_id, created_at DESC);
CREATE TABLE IF NOT EXISTS printer_health_checks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    printer_id INTEGER NOT NULL REFERENCES printers(id) ON DELETE CASCADE,
    is_reachable INTEGER NOT NULL CHECK(is_reachable IN (0,1)),
    response_ms INTEGER,
    error_message TEXT,
    checked_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS printer_health_recent_idx ON printer_health_checks(printer_id, checked_at DESC);
"""


def install_printer_schema(conn: sqlite3.Connection) -> None:
    """Install additive tables; safe for existing installations."""
    conn.executescript(SCHEMA_SQL)
    conn.commit()


def _json(value: Any) -> Any:
    if isinstance(value, str):
        try:
            return json.loads(value)
        except (TypeError, ValueError):
            return {}
    return value or {}


def _row(row: sqlite3.Row | None) -> dict[str, Any] | None:
    if row is None:
        return None
    result = dict(row)
    for key in ("driver_options", "capabilities", "payload"):
        if key in result and isinstance(result[key], str):
            result[key] = _json(result[key])
    return result


def _rows(rows: list[sqlite3.Row]) -> list[dict[str, Any]]:
    return [_row(row) for row in rows]


def register_printer_api(
    app,
    query: Callable[..., Any],
    execute: Callable[..., Any],
    connect_sqlite: Callable[[], sqlite3.Connection],
    login_required,
    admin_required,
) -> None:
    """Register printer management, assignment, and print queue endpoints."""

    def body() -> dict[str, Any]:
        return request.get_json(silent=True) or {}

    def validate_printer(data: dict[str, Any], partial: bool = False) -> str | None:
        if partial:
            return None
        missing = [key for key in ("profile_id", "name", "code", "connection") if not data.get(key)]
        if missing:
            return "الحقول المطلوبة: " + ", ".join(missing)
        if data.get("connection") not in CONNECTIONS:
            return "نوع الاتصال غير صالح"
        if data.get("connection") == "network" and (not data.get("address") or not data.get("port")):
            return "عنوان الشبكة والمنفذ مطلوبان للطابعة الشبكية"
        return None

    @app.get("/api/printer-profiles")
    @login_required
    def api_printer_profiles():
        return jsonify({"ok": True, "profiles": _rows(query("SELECT * FROM printer_profiles ORDER BY kind, name"))})

    @app.post("/api/printer-profiles")
    @admin_required
    def api_create_printer_profile():
        data = body()
        if data.get("kind") not in PRINTER_KINDS or data.get("payload_format") not in PAYLOAD_FORMATS or not data.get("name"):
            return jsonify({"ok": False, "error": "بيانات ملف الطابعة غير صالحة"}), 400
        now = NOW()
        profile_id = execute(
            """INSERT INTO printer_profiles(name,kind,paper_width_mm,label_width_mm,label_height_mm,dpi,payload_format,driver_options,created_at,updated_at)
               VALUES (?,?,?,?,?,?,?,?,?,?)""",
            (data["name"], data["kind"], data.get("paper_width_mm"), data.get("label_width_mm"), data.get("label_height_mm"), data.get("dpi"), data["payload_format"], json.dumps(data.get("driver_options") or {}, ensure_ascii=False), now, now),
        )
        return jsonify({"ok": True, "profile_id": profile_id}), 201

    @app.get("/api/printers")
    @login_required
    def api_list_printers():
        filters, args = [], []
        if request.args.get("kind"):
            filters.append("pp.kind=?"); args.append(request.args["kind"])
        if request.args.get("enabled") in ("0", "1"):
            filters.append("p.enabled=?"); args.append(int(request.args["enabled"]))
        where = ("WHERE " + " AND ".join(filters)) if filters else ""
        rows = query(f"""SELECT p.*, pp.name profile_name, pp.kind, pp.payload_format
                          FROM printers p JOIN printer_profiles pp ON pp.id=p.profile_id {where}
                          ORDER BY pp.kind, p.name""", args)
        return jsonify({"ok": True, "printers": _rows(rows)})

    @app.post("/api/printers")
    @admin_required
    def api_create_printer():
        data = body(); error = validate_printer(data)
        if error: return jsonify({"ok": False, "error": error}), 400
        profile = query("SELECT id FROM printer_profiles WHERE id=?", (data["profile_id"],), one=True)
        if not profile: return jsonify({"ok": False, "error": "ملف الطابعة غير موجود"}), 404
        now = NOW()
        try:
            printer_id = execute(
                """INSERT INTO printers(profile_id,name,code,connection,address,port,windows_printer_name,status,is_default,enabled,branch_id,workstation_id,capabilities,created_at,updated_at)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (data["profile_id"], data["name"], data["code"], data["connection"], data.get("address"), data.get("port"), data.get("windows_printer_name"), data.get("status", "inactive"), int(bool(data.get("is_default"))), int(data.get("enabled", True)), data.get("branch_id"), data.get("workstation_id"), json.dumps(data.get("capabilities") or {}, ensure_ascii=False), now, now),
            )
        except sqlite3.IntegrityError as exc:
            return jsonify({"ok": False, "error": str(exc)}), 409
        return jsonify({"ok": True, "printer_id": printer_id}), 201

    @app.route("/api/printers/<int:printer_id>", methods=["GET", "PATCH", "DELETE"])
    @login_required
    def api_printer_detail(printer_id: int):
        printer = query("SELECT p.*, pp.name profile_name, pp.kind, pp.payload_format FROM printers p JOIN printer_profiles pp ON pp.id=p.profile_id WHERE p.id=?", (printer_id,), one=True)
        if not printer: return jsonify({"ok": False, "error": "الطابعة غير موجودة"}), 404
        if request.method == "GET": return jsonify({"ok": True, "printer": _row(printer)})
        if session.get("role") != "مدير": return jsonify({"ok": False, "error": "يتطلب هذا الإجراء صلاحية المدير"}), 403
        if request.method == "DELETE":
            execute("UPDATE printers SET enabled=0,status='inactive',updated_at=? WHERE id=?", (NOW(), printer_id))
            return jsonify({"ok": True})
        data = body(); allowed = {"name", "connection", "address", "port", "windows_printer_name", "status", "is_default", "enabled", "branch_id", "workstation_id", "capabilities", "profile_id"}
        changes = [(key, data[key]) for key in allowed if key in data]
        if not changes: return jsonify({"ok": True, "printer": _row(printer)})
        assignments = ", ".join(f"{key}=?" for key, _ in changes) + ", updated_at=?"
        values = [json.dumps(value, ensure_ascii=False) if key == "capabilities" else value for key, value in changes] + [NOW(), printer_id]
        execute(f"UPDATE printers SET {assignments} WHERE id=?", values)
        return jsonify({"ok": True, "printer": _row(query("SELECT * FROM printers WHERE id=?", (printer_id,), one=True))})

    @app.post("/api/printers/<int:printer_id>/default")
    @admin_required
    def api_set_default_printer(printer_id: int):
        printer = query("SELECT profile_id FROM printers WHERE id=?", (printer_id,), one=True)
        if not printer: return jsonify({"ok": False, "error": "الطابعة غير موجودة"}), 404
        execute("UPDATE printers SET is_default=0, updated_at=? WHERE profile_id=?", (NOW(), printer["profile_id"]))
        execute("UPDATE printers SET is_default=1, enabled=1, status='active', updated_at=? WHERE id=?", (NOW(), printer_id))
        return jsonify({"ok": True})

    @app.post("/api/printer-assignments")
    @admin_required
    def api_create_assignment():
        data = body()
        if data.get("job_type") not in JOB_TYPES or not data.get("printer_id"):
            return jsonify({"ok": False, "error": "نوع المهمة والطابعة مطلوبان"}), 400
        try:
            assignment_id = execute(
                """INSERT INTO printer_assignments(printer_id,job_type,branch_id,workstation_id,priority,enabled,created_at,updated_at)
                   VALUES (?,?,?,?,?,?,?,?)""",
                (data["printer_id"], data["job_type"], data.get("branch_id"), data.get("workstation_id"), data.get("priority", 100), int(data.get("enabled", True)), NOW(), NOW()),
            )
        except sqlite3.IntegrityError as exc:
            return jsonify({"ok": False, "error": str(exc)}), 409
        return jsonify({"ok": True, "assignment_id": assignment_id}), 201

    @app.get("/api/printer-assignments")
    @login_required
    def api_list_assignments():
        rows = query("""SELECT pa.*, p.name printer_name, p.code printer_code, pp.kind
                        FROM printer_assignments pa JOIN printers p ON p.id=pa.printer_id
                        JOIN printer_profiles pp ON pp.id=p.profile_id ORDER BY pa.job_type, pa.priority""")
        return jsonify({"ok": True, "assignments": _rows(rows)})

    def select_printer(job_type: str, branch_id: Any = None, workstation_id: Any = None, explicit_id: Any = None):
        if explicit_id:
            row = query("SELECT p.*, pp.kind, pp.payload_format FROM printers p JOIN printer_profiles pp ON pp.id=p.profile_id WHERE p.id=? AND p.enabled=1 AND p.status='active'", (explicit_id,), one=True)
            if row: return row
            raise ValueError("EXPLICIT_PRINTER_UNAVAILABLE")
        row = query("""SELECT p.*, pp.kind, pp.payload_format,
                        (CASE WHEN pa.workstation_id IS NOT NULL THEN 2 ELSE 0 END + CASE WHEN pa.branch_id IS NOT NULL THEN 1 ELSE 0 END) specificity,
                        COALESCE(pa.priority, 100) assignment_priority
                     FROM printers p JOIN printer_profiles pp ON pp.id=p.profile_id
                     LEFT JOIN printer_assignments pa ON pa.printer_id=p.id AND pa.job_type=? AND pa.enabled=1
                       AND (pa.branch_id IS NULL OR pa.branch_id=?) AND (pa.workstation_id IS NULL OR pa.workstation_id=?)
                     WHERE p.enabled=1 AND p.status='active'
                     ORDER BY specificity DESC, assignment_priority ASC, p.is_default DESC, p.id ASC LIMIT 1""", (job_type, branch_id, workstation_id), one=True)
        if not row: raise ValueError(f"NO_ACTIVE_PRINTER_FOR_{job_type.upper()}")
        return row

    @app.post("/api/print-jobs")
    @login_required
    def api_create_print_job():
        data = body(); job_type = data.get("job_type"); fmt = data.get("payload_format")
        if job_type not in JOB_TYPES or fmt not in PAYLOAD_FORMATS:
            return jsonify({"ok": False, "error": "نوع المهمة أو تنسيق payload غير صالح"}), 400
        payload = data.get("payload"); raw = data.get("raw_payload_base64")
        if payload is None and not raw: return jsonify({"ok": False, "error": "payload مطلوب"}), 400
        key = request.headers.get("Idempotency-Key") or data.get("idempotency_key")
        if key:
            existing = query("SELECT * FROM print_jobs WHERE idempotency_key=?", (key,), one=True)
            if existing: return jsonify({"ok": True, "duplicate": True, "job": _row(existing)})
        try:
            printer = select_printer(job_type, data.get("branch_id"), data.get("workstation_id"), data.get("printer_id"))
        except ValueError as exc:
            return jsonify({"ok": False, "error": str(exc)}), 409
        try:
            job_id = execute(
                """INSERT INTO print_jobs(printer_id,job_type,status,payload_format,payload,raw_payload,copies,idempotency_key,source_type,source_id,requested_by,branch_id,workstation_id,queued_at,created_at,updated_at)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (printer["id"], job_type, "queued", fmt, json.dumps(payload, ensure_ascii=False) if payload is not None else None, base64.b64decode(raw) if raw else None, max(1, min(int(data.get("copies", 1)), 100)), key, data.get("source_type"), data.get("source_id"), session.get("user_id"), data.get("branch_id"), data.get("workstation_id"), NOW(), NOW(), NOW()),
            )
        except (ValueError, sqlite3.IntegrityError) as exc:
            return jsonify({"ok": False, "error": str(exc)}), 400
        return jsonify({"ok": True, "job_id": job_id, "printer_id": printer["id"], "status": "queued"}), 201

    @app.get("/api/print-jobs/<int:job_id>")
    @login_required
    def api_get_print_job(job_id: int):
        row = query("SELECT pj.*, p.name printer_name, p.code printer_code FROM print_jobs pj LEFT JOIN printers p ON p.id=pj.printer_id WHERE pj.id=?", (job_id,), one=True)
        if not row: return jsonify({"ok": False, "error": "مهمة الطباعة غير موجودة"}), 404
        return jsonify({"ok": True, "job": _row(row)})

    @app.post("/api/print-jobs/<int:job_id>/cancel")
    @login_required
    def api_cancel_print_job(job_id: int):
        execute("UPDATE print_jobs SET status='cancelled', updated_at=? WHERE id=? AND status='queued'", (NOW(), job_id))
        return jsonify({"ok": True})

    @app.post("/api/printers/<int:printer_id>/health")
    @login_required
    def api_printer_health(printer_id: int):
        # Actual TCP/USB probing belongs to the deployment-specific adapter.
        execute("INSERT INTO printer_health_checks(printer_id,is_reachable,error_message,checked_at) VALUES (?,?,?,?)", (printer_id, 0, "adapter_not_configured", NOW()))
        return jsonify({"ok": True, "reachable": False, "error": "adapter_not_configured"})
