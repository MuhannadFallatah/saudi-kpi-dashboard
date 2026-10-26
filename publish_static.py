# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path

from kpi_dashboard import MartService, RIYADH_TZ, WORKBOOK_PATH


ROOT = Path(__file__).resolve().parent
DOCS = ROOT / "docs"
DATA = DOCS / "data"
REFRESH_MINUTES = max(1, int(os.environ.get("KPI_REFRESH_MINUTES", "1440")))


def write_json(path: Path, payload: object) -> None:
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, allow_nan=False) + "\n",
        encoding="utf-8",
    )
    temporary.replace(path)


def main() -> None:
    DATA.mkdir(parents=True, exist_ok=True)
    service = MartService(WORKBOOK_PATH, REFRESH_MINUTES)
    result = service.refresh(live=True)
    if not result["ok"]:
        raise RuntimeError(result["message"])

    snapshot = service.snapshot()
    generated_at = datetime.now(RIYADH_TZ).isoformat(timespec="seconds")
    status = {
        **snapshot,
        "generatedAt": generated_at,
        "hostingMode": "GitHub Pages static site",
        "liveApi": False,
        "notice": (
            snapshot["notice"]
            + " This GitHub Pages deployment serves static JSON; it does not provide a running Python API."
        ),
    }

    write_json(DATA / "status.json", status)
    write_json(DATA / "facts.json", {"count": snapshot["factCount"], "data": service.facts_for()})

    excel_path = DATA / "Saudi_KPI_Extract.xlsx"
    excel_temp = excel_path.with_suffix(".xlsx.tmp")
    excel_temp.write_bytes(service.export_bytes())
    excel_temp.replace(excel_path)

    print(
        json.dumps(
            {
                "generatedAt": generated_at,
                "kpiCount": len(snapshot["kpis"]),
                "factCount": snapshot["factCount"],
                "greenCount": sum(kpi["refreshStatus"] == "green" for kpi in snapshot["kpis"]),
                "redCount": sum(kpi["refreshStatus"] == "red" for kpi in snapshot["kpis"]),
                "sourceWorkbook": str(WORKBOOK_PATH),
                "excelBytes": excel_path.stat().st_size,
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
