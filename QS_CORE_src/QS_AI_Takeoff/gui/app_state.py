"""Shared application state + background worker (keeps the GUI responsive)."""
from __future__ import annotations

import os
import queue
import threading
import traceback

from services.audit_service import AuditLog
from services.project_service import Project, ROOT


class AppState:
    def __init__(self):
        self.project: Project = Project.load()
        self.scan: dict | None = None
        self.scan_path: str = ""
        self.result = None
        self.excel_path: str = ""
        self.audit = AuditLog(os.path.join(ROOT, "logs", "audit.jsonl"))
        self.log_q: queue.Queue = queue.Queue()
        self._busy = threading.Lock()

    def log(self, msg: str):
        self.log_q.put(("log", msg))

    def run_bg(self, fn, on_done=None, label="task"):
        """Run fn() in a worker thread. Results are delivered through log_q to the Tk loop.
        AutoCAD COM calls are serialised by the lock so entities are never touched concurrently."""
        if not self._busy.acquire(blocking=False):
            self.log("Busy: another task is still running.")
            return False

        def work():
            try:
                self.log(f"[{label}] started")
                out = fn()
                self.log_q.put(("done", (on_done, out, None)))
            except Exception as e:  # noqa: BLE001
                self.log_q.put(("done", (on_done, None, f"{e}\n{traceback.format_exc()}")))
            finally:
                self._busy.release()
        threading.Thread(target=work, daemon=True).start()
        return True

    @property
    def output_dir(self) -> str:
        d = self.project.rules.get("output_dir", "output")
        return d if os.path.isabs(d) else os.path.join(ROOT, d)
