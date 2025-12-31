# -*- coding: utf-8 -*-

import json
import re
from datetime import datetime
from typing import Any, Iterable, List, Dict


ANSI_YELLOW = "\033[33m"
ANSI_RESET = "\033[0m"


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


def oid_from_uri(uri: str) -> str:
    m = re.search(r"[?&]oid=(\d+)", safe_str(uri))
    return m.group(1) if m else ""


def log(msg: str):
    print(msg)


def json_dumps_safe(obj: Any) -> str:
    try:
        return json.dumps(obj, ensure_ascii=False)
    except Exception:
        return "{}"


def apply_field(row: Dict[str, Any], key: str, value: str, *, overwrite: bool) -> bool:
    cur = safe_str(row.get(key))
    val = safe_str(value)
    if not val:
        return False
    if cur and not overwrite:
        return False
    row[key] = val
    return True
