# -*- coding: utf-8 -*-

from __future__ import annotations

import sys

from .config import Settings
from .pipeline import run
from .utils import log


def main() -> None:
    st = Settings.from_env()
    res = run(st)

    saved = res.request_meta.get("saved")
    rows = res.request_meta.get("rows")
    err = res.request_meta.get("error")

    log(f"saved: {saved} rows={rows}")

    if err:
        log(f"ERROR: {err}")
        sys.exit(1)

    sys.exit(0)


if __name__ == "__main__":
    main()
