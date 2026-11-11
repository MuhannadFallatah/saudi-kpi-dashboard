# -*- coding: utf-8 -*-
from __future__ import annotations

import gzip
import json
import os
import warnings
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from kpi_dashboard import MartService, RIYADH_TZ, WORKBOOK_PATH, download_bytes


ROOT = Path(__file__).resolve().parent
DOCS = ROOT / "docs"
DATA = DOCS / "data"
REFRESH_MINUTES = max(1, int(os.environ.get("KPI_REFRESH_MINUTES", "1440")))
PUBLISHED_STATUS_URL = "https://muhannadfallatah.github.io/saudi-kpi-dashboard/data/status.json"
PUBLISHED_FACTS_URL = "https://muhannadfallatah.github.io/saudi-kpi-dashboard/data/facts.json"


def write_json(path: Path, payload: object) -> None:
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, allow_nan=False) + "\n",
        encoding="utf-8",
    )
    temporary.replace(path)


def build_refresh_log(snapshot: dict[str, Any], generated_at: str) -> dict[str, Any]:
    kpis = snapshot.get("kpis", [])
    entries = [
        {
            "timestamp": kpi.get("lastChecked") or snapshot.get("lastRefresh") or generated_at,
            "kpiId": kpi.get("id"),
            "kpiName": kpi.get("name"),
            "status": kpi.get("refreshStatus", "red"),
            "label": kpi.get("refreshLabel", "Unknown"),
            "observationCount": kpi.get("observationCount", 0),
            "latestObservation": kpi.get("latestObservation"),
            "detail": kpi.get("refreshNote") or kpi.get("extractionStatus") or "No extraction details",
        }
        for kpi in kpis
    ]
    return {
        "generatedAt": generated_at,
        "lastRefresh": snapshot.get("lastRefresh") or generated_at,
        "lastError": snapshot.get("lastError"),
        "summary": {
            "kpiCount": len(entries),
            "successCount": sum(entry["status"] == "green" for entry in entries),
            "failedCount": sum(entry["status"] != "green" for entry in entries),
        },
        "entries": entries,
    }


def archive_daily_snapshot(
    archive_dir: Path,
    status: dict[str, Any],
    facts: list[dict[str, Any]],
    generated_at: str,
) -> dict[str, Any]:
    archive_date = datetime.fromisoformat(generated_at).date().isoformat()
    archive_dir.mkdir(parents=True, exist_ok=True)
    archive_path = archive_dir / f"{archive_date}.json.gz"
    index_path = archive_dir / "index.json"
    created = not archive_path.exists()

    if created:
        payload = {
            "archiveDate": archive_date,
            "archivedAt": generated_at,
            "factCount": len(facts),
            "status": status,
            "facts": facts,
        }
        temporary = archive_path.with_suffix(archive_path.suffix + ".tmp")
        temporary.write_bytes(
            gzip.compress(
                json.dumps(payload, ensure_ascii=False, indent=2, allow_nan=False).encode("utf-8"),
                compresslevel=9,
                mtime=0,
            )
        )
        temporary.replace(archive_path)

    if index_path.exists():
        index = json.loads(index_path.read_text(encoding="utf-8"))
        if not isinstance(index, dict) or not isinstance(index.get("snapshots"), list):
            raise ValueError(f"Invalid archive index: {index_path}")
    else:
        index = {"snapshots": []}

    snapshots = {item["date"]: item for item in index["snapshots"]}
    if archive_date not in snapshots:
        if not created:
            with gzip.open(archive_path, "rt", encoding="utf-8") as archived_file:
                archived = json.load(archived_file)
        else:
            archived = payload
        archived_status = archived["status"]
        snapshots[archive_date] = {
            "date": archive_date,
            "archivedAt": archived["archivedAt"],
            "factCount": archived["factCount"],
            "kpiCount": len(archived_status.get("kpis", [])),
            "url": f"{archive_date}.json.gz",
        }

    index["snapshots"] = sorted(snapshots.values(), key=lambda item: item["date"], reverse=True)
    write_json(index_path, index)
    return {
        "date": archive_date,
        "created": created,
        "factCount": snapshots[archive_date]["factCount"],
    }


def load_pre_refresh_snapshot(
    data_dir: Path,
    fallback_status: dict[str, Any],
    fallback_facts: list[dict[str, Any]],
    fetch: Callable[[str], bytes] = download_bytes,
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    try:
        status = json.loads(fetch(PUBLISHED_STATUS_URL))
        facts_payload = json.loads(fetch(PUBLISHED_FACTS_URL))
        facts = facts_payload.get("data") if isinstance(facts_payload, dict) else None
        if not isinstance(status, dict) or not isinstance(facts, list):
            raise ValueError("Published pre-refresh KPI status or facts have an invalid format.")
        return status, facts
    except (OSError, ValueError) as error:
        warnings.warn(
            f"Could not load the current published dataset for pre-refresh archiving: {error}",
            RuntimeWarning,
        )

    status_path = data_dir / "status.json"
    facts_path = data_dir / "facts.json"
    if status_path.exists() and facts_path.exists():
        status = json.loads(status_path.read_text(encoding="utf-8"))
        facts_payload = json.loads(facts_path.read_text(encoding="utf-8"))
        facts = facts_payload.get("data") if isinstance(facts_payload, dict) else None
        if not isinstance(status, dict) or not isinstance(facts, list):
            raise ValueError(f"Invalid previously published KPI data in {data_dir}.")
        return status, facts

    return fallback_status, fallback_facts


def main() -> None:
    DATA.mkdir(parents=True, exist_ok=True)
    service = MartService(WORKBOOK_PATH, REFRESH_MINUTES)
    generated_at = datetime.now(RIYADH_TZ).isoformat(timespec="seconds")
    archive_date = datetime.fromisoformat(generated_at).date().isoformat()
    pre_refresh_status, pre_refresh_facts = load_pre_refresh_snapshot(
        DATA,
        service.snapshot(),
        service.facts_for(),
    )
    archive = archive_daily_snapshot(
        DOCS / "archive",
        pre_refresh_status,
        pre_refresh_facts,
        generated_at,
    )

    result = service.refresh(live=True)
    if not result["ok"]:
        raise RuntimeError(result["message"])

    snapshot = service.snapshot()
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
    write_json(DATA / "refresh-logs.json", build_refresh_log(status, generated_at))
    write_json(DATA / "facts.json", {"count": snapshot["factCount"], "data": service.facts_for()})

    kpi_dir = DATA / "kpi"
    kpi_dir.mkdir(exist_ok=True)
    for old in kpi_dir.glob("*.json"):
        old.unlink()
    for kpi in snapshot["kpis"]:
        rows = service.facts_for(kpi["id"])
        write_json(kpi_dir / f"{kpi['id']}.json", {"kpi": kpi, "count": len(rows), "data": rows})
    write_json(
        kpi_dir / "index.json",
        [
            {
                "id": kpi["id"],
                "name": kpi["name"],
                "refreshStatus": kpi["refreshStatus"],
                "observationCount": kpi["observationCount"],
                "url": f"{kpi['id']}.json",
            }
            for kpi in snapshot["kpis"]
        ],
    )

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
                "archiveDate": archive["date"],
                "archiveCreated": archive["created"],
                "archiveFactCount": archive["factCount"],
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
