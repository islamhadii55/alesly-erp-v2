"""Local USB printer discovery and registration helpers.

USB access is intentionally performed by the local Python process, not by the
browser.  The module works with Linux device nodes/CUPS and optionally with
Windows Print Spooler when pywin32 is installed.
"""
from __future__ import annotations

import hashlib
import os
import platform
import re
import subprocess
from datetime import datetime
from typing import Any


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _code(value: str) -> str:
    return "USB-" + hashlib.sha1(value.encode("utf-8", "ignore")).hexdigest()[:10].upper()


def discover_usb_printers() -> list[dict[str, Any]]:
    """Return locally visible USB printers without raising on unsupported hosts."""
    found: list[dict[str, Any]] = []
    system = platform.system().lower()

    if system == "windows":
        try:
            import win32print  # type: ignore

            flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            for item in win32print.EnumPrinters(flags, None, 2):
                name = str(item[2] or "").strip()
                port = str(item[5] or "").strip()
                if name and ("usb" in port.lower() or "usb" in name.lower()):
                    found.append({"key": f"windows:{name}", "name": name, "address": port, "windows_printer_name": name})
        except (ImportError, OSError):
            pass
        return found

    # CUPS exposes USB printers as usb:// URIs.  lpstat is available on most
    # Linux desktop installations and is preferable to requiring libusb/root.
    try:
        result = subprocess.run(["lpstat", "-v"], capture_output=True, text=True, timeout=3, check=False)
        for line in result.stdout.splitlines():
            match = re.match(r"device for (.+?):\s*(usb://.+)$", line.strip(), re.I)
            if match:
                name, uri = match.group(1).strip(), match.group(2).strip()
                found.append({"key": uri, "name": name, "address": uri, "windows_printer_name": name})
    except (FileNotFoundError, OSError, subprocess.TimeoutExpired):
        pass

    # Direct USB printer nodes are used by lightweight POS deployments.
    for path in sorted(["/dev/usb/lp0", "/dev/usb/lp1", "/dev/usb/lp2", "/dev/usb/lp3"]):
        if os.path.exists(path):
            found.append({"key": path, "name": f"USB Printer {os.path.basename(path)}", "address": path})

    unique: dict[str, dict[str, Any]] = {}
    for printer in found:
        unique[printer["key"]] = printer
    return list(unique.values())


def sync_usb_printers(connect_sqlite) -> list[dict[str, Any]]:
    """Register discovered USB printers and mark unplugged ones inactive."""
    discovered = discover_usb_printers()
    conn = connect_sqlite()
    conn.row_factory = __import__("sqlite3").Row
    try:
        now = _now()
        profile = conn.execute("SELECT id FROM printer_profiles WHERE kind='thermal' ORDER BY id LIMIT 1").fetchone()
        if not profile:
            conn.execute(
                "INSERT INTO printer_profiles(name,kind,paper_width_mm,payload_format,driver_options,created_at,updated_at) VALUES (?,?,?,?,?,?,?)",
                ("طابعة USB تلقائية", "thermal", 80, "esc_pos", "{}", now, now),
            )
            profile = conn.execute("SELECT id FROM printer_profiles WHERE kind='thermal' ORDER BY id LIMIT 1").fetchone()
        profile_id = profile["id"]
        seen_codes: list[str] = []
        for item in discovered:
            code = _code(item["key"])
            seen_codes.append(code)
            existing = conn.execute("SELECT id FROM printers WHERE code=?", (code,)).fetchone()
            if existing:
                conn.execute(
                    "UPDATE printers SET name=?,address=?,windows_printer_name=?,status='active',enabled=1,last_seen_at=?,last_error=NULL,updated_at=?,capabilities=? WHERE id=?",
                    (item["name"], item.get("address"), item.get("windows_printer_name"), now, now, '{"auto_discovered":true,"connection":"usb"}', existing["id"]),
                )
            else:
                conn.execute(
                    "INSERT INTO printers(profile_id,name,code,connection,address,port,windows_printer_name,status,is_default,enabled,capabilities,created_at,updated_at,last_seen_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    (profile_id, item["name"], code, "usb", item.get("address"), None, item.get("windows_printer_name"), "active", 0, 1, '{"auto_discovered":true,"connection":"usb"}', now, now, now),
                )
        if seen_codes:
            placeholders = ",".join("?" for _ in seen_codes)
            conn.execute(f"UPDATE printers SET status='inactive',updated_at=? WHERE connection='usb' AND enabled=1 AND code NOT IN ({placeholders})", [now, *seen_codes])
            default_exists = conn.execute("SELECT id FROM printers WHERE profile_id=? AND is_default=1 AND enabled=1 LIMIT 1", (profile_id,)).fetchone()
            if not default_exists:
                conn.execute("UPDATE printers SET is_default=1 WHERE code=?", (seen_codes[0],))
        else:
            conn.execute("UPDATE printers SET status='inactive',updated_at=? WHERE connection='usb' AND enabled=1", (now,))
        conn.commit()
        return discovered
    finally:
        conn.close()
