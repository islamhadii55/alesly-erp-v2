"""Deployment adapters for physical printers."""
from __future__ import annotations

import os
import platform
import socket
import subprocess
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
        payload = b"\x1b@\x1ba\x01" + name + b"\n"
        payload += b"\x1dH\x02\x1dw\x02\x1dh\x3c\x1dk\x49" + bytes([len(code) + 2]) + b"\x7b\x42" + code
        payload += b"\n\n\x1dV\x00"
        for _ in range(max(1, min(int(copies), 100))):
            self.print_bytes(config, payload)


@dataclass
class UsbPrinterAdapter:
    """Send raw ESC/POS to a local USB printer through the OS print layer."""
    timeout: float = 5.0

    def test_connection(self, config: dict[str, Any]) -> dict[str, Any]:
        address = str(config.get("address") or "")
        name = str(config.get("windows_printer_name") or "")
        if address.startswith("/dev/"):
            reachable = os.path.exists(address)
            return {"reachable": reachable, "error": None if reachable else "USB_DEVICE_NOT_FOUND"}
        if platform.system().lower() == "windows" and name:
            try:
                import win32print  # type: ignore
                handle = win32print.OpenPrinter(name)
                win32print.ClosePrinter(handle)
                return {"reachable": True}
            except (ImportError, OSError):
                return {"reachable": False, "error": "WINDOWS_PRINTER_UNAVAILABLE"}
        if name:
            try:
                result = subprocess.run(["lpstat", "-p", name], capture_output=True, text=True, timeout=self.timeout, check=False)
                return {"reachable": result.returncode == 0, "error": None if result.returncode == 0 else (result.stderr.strip() or "CUPS_PRINTER_UNAVAILABLE")}
            except (FileNotFoundError, OSError, subprocess.TimeoutExpired):
                pass
        return {"reachable": False, "error": "USB_PRINTER_ADAPTER_UNAVAILABLE"}

    def print_bytes(self, config: dict[str, Any], payload: bytes) -> None:
        address = str(config.get("address") or "")
        name = str(config.get("windows_printer_name") or "")
        if address.startswith("/dev/"):
            with open(address, "wb", buffering=0) as device:
                device.write(payload)
            return
        if platform.system().lower() == "windows" and name:
            import win32print  # type: ignore
            handle = win32print.OpenPrinter(name)
            try:
                win32print.StartDocPrinter(handle, 1, ("ALES ERP", None, "RAW"))
                win32print.StartPagePrinter(handle)
                win32print.WritePrinter(handle, payload)
                win32print.EndPagePrinter(handle)
                win32print.EndDocPrinter(handle)
            finally:
                win32print.ClosePrinter(handle)
            return
        if name:
            result = subprocess.run(["lp", "-d", name, "-o", "raw"], input=payload, capture_output=True, timeout=self.timeout, check=False)
            if result.returncode == 0:
                return
            raise RuntimeError(result.stderr.decode("utf-8", "replace") or "CUPS_PRINT_FAILED")
        raise RuntimeError("USB_PRINTER_ENDPOINT_REQUIRED")

    def print_receipt(self, config: dict[str, Any], text: str, copies: int = 1) -> None:
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
        options = config.get("driver_options") or {}
        encoding = str(options.get("encoding") or "cp864")
        name = product_name.encode(encoding, errors="replace")
        code = value.encode("ascii", errors="ignore")
        if not code:
            raise ValueError("BARCODE_VALUE_REQUIRED")
        payload = b"\x1b@\x1ba\x01" + name + b"\n"
        payload += b"\x1dH\x02\x1dw\x02\x1dh\x3c\x1dk\x49" + bytes([len(code) + 2]) + b"\x7b\x42" + code + b"\n\n\x1dV\x00"
        for _ in range(max(1, min(int(copies), 100))):
            self.print_bytes(config, payload)
