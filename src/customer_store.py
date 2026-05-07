"""Local JSON persistence for customer discount history (POC)."""

from __future__ import annotations

import json
from pathlib import Path

_STORE_PATH = Path(__file__).resolve().parent.parent / "data" / "customers.json"


def _load() -> dict[str, dict]:
    if not _STORE_PATH.exists():
        return {}
    try:
        data = json.loads(_STORE_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}
    return data if isinstance(data, dict) else {}


def _dump(data: dict[str, dict]) -> None:
    _STORE_PATH.parent.mkdir(parents=True, exist_ok=True)
    _STORE_PATH.write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def get_discount(customer_id: str) -> float | None:
    row = _load().get(customer_id.strip())
    if not row:
        return None
    pct = row.get("discount_percent")
    if pct is None:
        return None
    return float(pct)


def save_discount(customer_id: str, discount_percent: float) -> None:
    cid = customer_id.strip()
    data = _load()
    data[cid] = {"discount_percent": float(discount_percent)}
    _dump(data)
