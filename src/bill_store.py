"""Local persistence for generated báo giá files (POC)."""

from __future__ import annotations

import json
import uuid
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

_BILLS_DIR = Path(__file__).resolve().parent.parent / "data" / "bills"
_INDEX_PATH = Path(__file__).resolve().parent.parent / "data" / "bills_index.json"


@dataclass
class BillRecord:
    id: str
    customer_id: str
    created_at: str
    source_filename: str
    output_filename: str
    file_path: str
    discount_percent: float
    delivery_days: int

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> BillRecord:
        return cls(
            id=str(data["id"]),
            customer_id=str(data["customer_id"]),
            created_at=str(data["created_at"]),
            source_filename=str(data.get("source_filename", "")),
            output_filename=str(data.get("output_filename", "")),
            file_path=str(data["file_path"]),
            discount_percent=float(data.get("discount_percent", 0)),
            delivery_days=int(data.get("delivery_days", 0)),
        )


def _load_index() -> list[dict[str, Any]]:
    if not _INDEX_PATH.exists():
        return []
    try:
        data = json.loads(_INDEX_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return []
    return data if isinstance(data, list) else []


def _save_index(records: list[dict[str, Any]]) -> None:
    _INDEX_PATH.parent.mkdir(parents=True, exist_ok=True)
    _INDEX_PATH.write_text(
        json.dumps(records, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def save_bill(
    customer_id: str,
    output_bytes: bytes,
    *,
    source_filename: str,
    output_filename: str,
    discount_percent: float,
    delivery_days: int,
) -> str:
    cid = customer_id.strip()
    bill_id = uuid.uuid4().hex[:12]
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S")
    safe_name = output_filename.replace("/", "_").replace("\\", "_")
    rel_dir = Path("bills") / cid
    out_dir = _BILLS_DIR.parent / rel_dir
    out_dir.mkdir(parents=True, exist_ok=True)
    file_name = f"{stamp}_{safe_name}"
    file_path = out_dir / file_name
    file_path.write_bytes(output_bytes)

    created_at = datetime.now(timezone.utc).isoformat()
    record = {
        "id": bill_id,
        "customer_id": cid,
        "created_at": created_at,
        "source_filename": source_filename,
        "output_filename": output_filename,
        "file_path": str(rel_dir / file_name),
        "discount_percent": float(discount_percent),
        "delivery_days": int(delivery_days),
    }
    records = _load_index()
    records.append(record)
    _save_index(records)
    return bill_id


def list_bills(customer_id: str, limit: int = 50) -> list[BillRecord]:
    cid = customer_id.strip().lower()
    matches = [
        BillRecord.from_dict(r)
        for r in _load_index()
        if str(r.get("customer_id", "")).strip().lower() == cid
    ]
    matches.sort(key=lambda b: b.created_at, reverse=True)
    return matches[:limit]


def load_bill(bill_id: str) -> bytes | None:
    bid = bill_id.strip()
    for record in _load_index():
        if str(record.get("id")) != bid:
            continue
        rel = Path(str(record["file_path"]))
        path = _BILLS_DIR.parent / rel
        if path.is_file():
            return path.read_bytes()
    return None
