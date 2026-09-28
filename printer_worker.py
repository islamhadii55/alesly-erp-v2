"""Persistent print queue worker for thermal network and USB printers."""
from __future__ import annotations

import argparse
import json
import logging
import signal
import time
from datetime import datetime
from typing import Any

from printer_adapters import ThermalTCPAdapter, UsbPrinterAdapter
from printer_discovery import sync_usb_printers

LOG = logging.getLogger("printer-worker")


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


class PrintQueueWorker:
    def __init__(self, connect_sqlite, poll_seconds: float = 1.0, max_attempts: int = 3, retry_delay: float = 2.0, adapter=None, usb_scan_seconds: float = 3.0):
        self.connect_sqlite = connect_sqlite
        self.poll_seconds = max(0.1, poll_seconds)
        self.max_attempts = max(1, max_attempts)
        self.retry_delay = max(0.0, retry_delay)
        self.adapter = adapter or ThermalTCPAdapter()
        self.usb_adapter = UsbPrinterAdapter()
        self.usb_scan_seconds = max(1.0, usb_scan_seconds)
        self.last_usb_scan = 0.0
        self.running = True

    def stop(self, *_args) -> None:
        self.running = False

    def claim_next(self) -> dict[str, Any] | None:
        conn = self.connect_sqlite()
        conn.row_factory = __import__("sqlite3").Row
        try:
            conn.execute("BEGIN IMMEDIATE")
            job = conn.execute("""SELECT pj.*, p.address, p.port, p.connection, p.status printer_status,
                                        p.name printer_name, p.windows_printer_name,
                                        pp.kind, pp.payload_format profile_payload_format, pp.driver_options
                                 FROM print_jobs pj
                                 LEFT JOIN printers p ON p.id=pj.printer_id
                                 LEFT JOIN printer_profiles pp ON pp.id=p.profile_id
                                 WHERE pj.status='queued'
                                 ORDER BY pj.queued_at, pj.id LIMIT 1""").fetchone()
            if not job:
                conn.rollback()
                return None
            stamp = now()
            conn.execute("UPDATE print_jobs SET status='printing', attempts=attempts+1, started_at=?, updated_at=? WHERE id=? AND status='queued'", (stamp, stamp, job["id"]))
            conn.commit()
            result = dict(job)
            result["attempts"] = int(job["attempts"]) + 1
            return result
        finally:
            conn.close()

    def _config(self, job: dict[str, Any]) -> dict[str, Any]:
        try:
            options = json.loads(job.get("driver_options") or "{}")
        except (TypeError, ValueError):
            options = {}
        return {
            "address": job.get("address"),
            "port": job.get("port") or 9100,
            "windows_printer_name": job.get("windows_printer_name") or job.get("printer_name"),
            "driver_options": options,
        }

    def deliver(self, job: dict[str, Any]) -> None:
        if job.get("kind") != "thermal" or job.get("connection") not in ("network", "usb"):
            raise RuntimeError("UNSUPPORTED_PRINTER_ADAPTER_OR_CONNECTION")
        config = self._config(job)
        adapter = self.usb_adapter if job.get("connection") == "usb" else self.adapter
        copies = int(job.get("copies") or 1)
        if job.get("raw_payload") is not None:
            adapter.print_bytes(config, bytes(job["raw_payload"]))
            return
        payload = json.loads(job.get("payload") or "{}")
        if job.get("job_type") == "barcode_label":
            adapter.print_barcode_label(config, str(payload.get("value") or ""), str(payload.get("product_name") or ""), copies)
        elif job.get("payload_format") == "esc_pos":
            text = payload.get("text") or payload.get("content") or json.dumps(payload, ensure_ascii=False)
            adapter.print_receipt(config, str(text), copies)
        else:
            raise RuntimeError(f"UNSUPPORTED_PAYLOAD_FORMAT_{job.get('payload_format')}")

    def mark_completed(self, job_id: int) -> None:
        conn = self.connect_sqlite()
        try:
            stamp = now()
            conn.execute("UPDATE print_jobs SET status='completed', completed_at=?, error_message=NULL, updated_at=? WHERE id=?", (stamp, stamp, job_id))
            conn.execute("UPDATE printers SET status='active', last_seen_at=?, last_error=NULL, updated_at=? WHERE id=(SELECT printer_id FROM print_jobs WHERE id=?)", (stamp, stamp, job_id))
            conn.execute("INSERT INTO printer_health_checks(printer_id,is_reachable,response_ms,error_message,checked_at) SELECT printer_id,1,NULL,NULL,? FROM print_jobs WHERE id=? AND printer_id IS NOT NULL", (stamp, job_id))
            conn.commit()
        finally:
            conn.close()

    def mark_failed_or_retry(self, job: dict[str, Any], error: Exception) -> str:
        conn = self.connect_sqlite()
        try:
            stamp = now()
            message = str(error)[:1000]
            if int(job["attempts"]) < self.max_attempts:
                conn.execute("UPDATE print_jobs SET status='queued', error_message=?, updated_at=? WHERE id=?", (f"محاولة {job['attempts']}/{self.max_attempts}: {message}", stamp, job["id"]))
                result = "retry"
            else:
                conn.execute("UPDATE print_jobs SET status='failed', error_message=?, updated_at=? WHERE id=?", (f"فشل بعد {job['attempts']} محاولات: {message}", stamp, job["id"]))
                result = "failed"
            conn.execute("UPDATE printers SET status='error', last_error=?, updated_at=? WHERE id=?", (message, stamp, job.get("printer_id")))
            if job.get("printer_id"):
                conn.execute("INSERT INTO printer_health_checks(printer_id,is_reachable,response_ms,error_message,checked_at) VALUES (?,?,NULL,?,?)", (job.get("printer_id"), 0, message, stamp))
            conn.commit()
            return result
        finally:
            conn.close()

    def process_once(self) -> str:
        job = self.claim_next()
        if not job:
            return "idle"
        try:
            self.deliver(job)
            self.mark_completed(job["id"])
            LOG.info("job=%s status=completed printer=%s", job["id"], job.get("printer_id"))
            return "completed"
        except Exception as exc:
            result = self.mark_failed_or_retry(job, exc)
            LOG.error("job=%s status=%s printer=%s error=%s", job["id"], result, job.get("printer_id"), exc)
            if result == "retry" and self.retry_delay:
                time.sleep(self.retry_delay)
            return result

    def run_forever(self) -> None:
        LOG.info("print worker started poll_seconds=%s usb_scan_seconds=%s", self.poll_seconds, self.usb_scan_seconds)
        while self.running:
            if time.monotonic() - self.last_usb_scan >= self.usb_scan_seconds:
                try:
                    discovered = sync_usb_printers(self.connect_sqlite)
                    if discovered:
                        LOG.info("usb printers discovered=%s", len(discovered))
                except Exception as exc:
                    LOG.warning("usb discovery failed: %s", exc)
                self.last_usb_scan = time.monotonic()
            if self.process_once() == "idle":
                time.sleep(self.poll_seconds)
        LOG.info("print worker stopped")


def main() -> None:
    parser = argparse.ArgumentParser(description="Run the thermal printer queue worker")
    parser.add_argument("--poll-seconds", type=float, default=1.0)
    parser.add_argument("--max-attempts", type=int, default=3)
    parser.add_argument("--retry-delay", type=float, default=2.0)
    args = parser.parse_args()
    import app
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s %(message)s")
    worker = PrintQueueWorker(app.connect_sqlite, args.poll_seconds, args.max_attempts, args.retry_delay)
    signal.signal(signal.SIGINT, worker.stop)
    signal.signal(signal.SIGTERM, worker.stop)
    worker.run_forever()


if __name__ == "__main__":
    main()
