"""Load / save / validate project configuration JSON."""
from __future__ import annotations

import json
import os
from copy import deepcopy

from engine.validation import validate_project

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_RULES = os.path.join(ROOT, "config", "project_rules.json")
DEFAULT_LAYERS = os.path.join(ROOT, "config", "layer_mapping.json")


def load_json(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_json(path: str, data: dict) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


class Project:
    def __init__(self, rules: dict, layers: dict, rules_path: str = "", layers_path: str = ""):
        self.rules = rules
        self.layers = layers
        self.rules_path = rules_path
        self.layers_path = layers_path

    @classmethod
    def load(cls, rules_path: str = DEFAULT_RULES, layers_path: str = DEFAULT_LAYERS) -> "Project":
        return cls(load_json(rules_path), load_json(layers_path), rules_path, layers_path)

    def save(self, rules_path: str | None = None, layers_path: str | None = None) -> None:
        self.rules_path = rules_path or self.rules_path
        self.layers_path = layers_path or self.layers_path
        save_json(self.rules_path, self.rules)
        save_json(self.layers_path, self.layers)

    def validate(self) -> list:
        return validate_project(self.rules)

    @property
    def unit_to_m(self) -> float:
        return float(self.rules["unit_to_m"])

    def detection_cfg(self) -> dict:
        """Detection tolerances are stored in metres in JSON; convert to drawing units."""
        d = deepcopy(self.rules["detection"])
        k = 1.0 / self.unit_to_m
        for key in ("min_thickness", "max_thickness", "min_wall_length", "max_endpoint_mismatch",
                    "gap_tol", "dup_tol", "opening_search_dist", "max_bridge_gap",
                    "bridge_thickness_tol"):
            if key in d:
                d[key] = float(d[key]) * k
        return d
