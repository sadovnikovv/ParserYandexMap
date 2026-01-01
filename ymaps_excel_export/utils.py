# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import math
import os
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, Iterable, List, Optional, Tuple

ANSI_YELLOW = "\033[33m"
ANSI_RESET = "\033[0m"


def log(msg: str) -> None:
    print(msg, flush=True)


def now_str_for_filename() -> str:
    return datetime.now().strftime("%Y-%m-%d_%H-%M-%S")


def now_iso_local() -> str:
    return datetime.now().astimezone().replace(microsecond=0).isoformat()


def safe_str(x: Any) -> str:
    if x is None:
        return ""
    if isinstance(x, str):
        return x.strip()
    return str(x).strip()


def json_dumps_safe(obj: Any) -> str:
    try:
        return json.dumps(obj, ensure_ascii=False)
    except Exception:
        return "{}"


def dedup_keep_order(items: Iterable[str]) -> List[str]:
    seen = set()
    out: List[str] = []
    for x in items:
        x = safe_str(x)
        if not x or x in seen:
            continue
        seen.add(x)
        out.append(x)
    return out


def pick_n(items: List[str], n: int) -> List[str]:
    items = items[:n]
    if len(items) < n:
        items += [""] * (n - len(items))
    return items


def safe_join(items: Iterable[str], sep: str = ", ") -> str:
    return sep.join(dedup_keep_order(items))


def oid_from_uri(uri: str) -> str:
    m = re.search(r"[?&]oid=(\d+)", safe_str(uri))
    return m.group(1) if m else ""


def apply_field(dst: Dict[str, Any], key: str, value: str, *, overwrite: bool) -> bool:
    cur = safe_str(dst.get(key))
    val = safe_str(value)
    if not val:
        return False
    if cur and not overwrite:
        return False
    dst[key] = val
    return True


def bbox_from_center_diameter_km(center_lon: float, center_lat: float, diameter_km: float) -> str:
    """
    bbox формата: "lon1,lat1~lon2,lat2"
    """
    if diameter_km <= 0:
        raise ValueError("DIAMETER_KM must be > 0")

    radius_km = diameter_km / 2.0
    km_per_deg_lat = 110.574
    km_per_deg_lon = 111.320 * math.cos(math.radians(center_lat))
    if abs(km_per_deg_lon) < 1e-9:
        km_per_deg_lon = 1e-9

    dlat = radius_km / km_per_deg_lat
    dlon = radius_km / km_per_deg_lon

    lon1 = max(-180.0, min(180.0, center_lon - dlon))
    lon2 = max(-180.0, min(180.0, center_lon + dlon))
    lat1 = max(-90.0, min(90.0, center_lat - dlat))
    lat2 = max(-90.0, min(90.0, center_lat + dlat))

    return f"{lon1:.6f},{lat1:.6f}~{lon2:.6f},{lat2:.6f}"


def env_str(name: str, default: str = "") -> str:
    v = os.getenv(name)
    return safe_str(v) if v is not None else safe_str(default)


def env_int(name: str, default: int) -> int:
    v = env_str(name, "")
    return int(v) if v else default


def env_float(name: str, default: float) -> float:
    v = env_str(name, "")
    return float(v) if v else default


def env_bool01(name: str, default: bool) -> bool:
    v = env_str(name, "")
    if not v:
        return default
    v = v.lower()
    if v in ("1", "true", "yes", "on"):
        return True
    if v in ("0", "false", "no", "off"):
        return False
    return default
