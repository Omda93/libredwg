"""Append-only JSON-lines audit log."""
from __future__ import annotations

import json
import os
from datetime import datetime


class AuditLog:
    def __init__(self, path: str):
        self.path = path
        os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
        self.entries: list = []
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line:
                        try:
                            self.entries.append(json.loads(line))
                        except json.JSONDecodeError:
                            pass

    def record(self, action: str, target: str = "", detail: str = "", user: str = "") -> dict:
        e = {"timestamp": datetime.now().isoformat(timespec="seconds"), "action": action,
             "target": target, "detail": detail, "user": user or os.environ.get("USERNAME", "")}
        self.entries.append(e)
        with open(self.path, "a", encoding="utf-8") as f:
            f.write(json.dumps(e, ensure_ascii=False) + "\n")
        return e
