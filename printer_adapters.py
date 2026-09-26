"""Deployment adapters for physical printers.

The thermal adapter sends ESC/POS bytes over a local TCP socket (usually port 9100).
It never opens a connection until print() or test_connection() is explicitly called.
"""
from __future__ import annotations

import socket
from dataclasses import dataclass
from typing import Any


@dataclass
class ThermalTCPAdapter:
    timeout: float = 3.0

    def _endpoint(self, config: dict[str, Any]) -> tuple[str, int]:
        host = str(config.get("address") or "").strip()
        port = int(config.get("port") or 9100)
        if not host:
            raise ValueError("THERMAL_PRINTER_ADDRESS_REQUIRED")
        if not 1 <= port <= 65535:
            raise ValueError("THERMAL_PRINTER_PORT_INVALID")
        return host, port

    def test_connection(self, config: dict[str, Any]) -> dict[str, Any]:
        host, port = self._endpoint(config)
        started = __import__("time").perf_counter()
        try:
            with socket.create_connection((host, port), timeout=self.timeout):
                elapsed = round((__import__("time").perf_counter() - started) * 1000)
                return {"reachable": True, "response_ms": elapsed}
        except (OSError, ValueError) as exc:
            return {"reachable": False, "error": str(exc)}

    def print_bytes(self, config: dict[str, Any], payload: bytes) -> None:
        host, port = self._endpoint(config)
        with socket.create_connection((host, port), timeout=self.timeout) as printer:
            printer.sendall(payload)

    def print_receipt(self, config: dict[str, Any], text: str, copies: int = 1) -> None:
        """Print UTF-8/driver-encoded text. Arabic glyph shaping is driver-dependent."""
        options = config.get("driver_options") or {}
        encoding = str(options.get("encoding") or "cp864")
        try:
            encoded = text.encode(encoding, errors="replace")
        except LookupError:
            encoded = text.encode("utf-8", errors="replace")
        payload = b"\x1b@" + b"\x1b\x61\x01" + encoded + b"\n\n\x1dV\x00"
        for _ in range(max(1, min(int(copies), 100))):
            self.print_bytes(config, payload)

    def print_barcode_label(self, config: dict[str, Any], value: str, product_name: str = "", copies: int = 1) -> None:
        """Print a Code128 label using ESC/POS barcode commands."""
        options = config.get("driver_options") or {}
        encoding = str(options.get("encoding") or "cp864")
        try:
            name = product_name.encode(encoding, errors="replace")
            code = value.encode("ascii", errors="ignore")
        except LookupError:
            name = product_name.encode("utf-8", errors="replace")
            code = value.encode("ascii", errors="ignore")
        if not code:
            raise ValueError("BARCODE_VALUE_REQUIRED")
        # Center, human-readable text, height 60, width 2, Code128 subset B.
        payload = b"\x1b@\x1ba\x01" + name + b"\n"
        payload += b"\x1dH\x02\x1dw\x02\x1dh\x3c"
        payload += b"\x1dk\x49" + bytes([len(code) + 2]) + b"\x7b\x42" + code
        payload += b"\n\n\x1dV\x00"
        for _ in range(max(1, min(int(copies), 100))):
            self.print_bytes(config, payload)
