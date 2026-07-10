import json
import os
import threading
import uuid
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from typing import Any


PROJECT_DIR = Path(__file__).resolve().parent
DEFAULT_DATA_DIR = PROJECT_DIR / "outputs" / "online_bridge"
BRIDGE_STORE_VERSION = 1
_LOCK = threading.Lock()


def clean_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def data_dir() -> Path:
    raw = clean_text(os.environ.get("CADASTRO_DATA_DIR"))
    path = Path(raw) if raw else DEFAULT_DATA_DIR
    path = path.resolve()
    try:
        path.mkdir(parents=True, exist_ok=True)
        probe = path / ".write_test"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink(missing_ok=True)
        return path
    except OSError:
        fallback = Path(os.environ.get("TMPDIR", "/tmp")) / "modulo-cadastro-data"
        fallback.mkdir(parents=True, exist_ok=True)
        return fallback.resolve()


def store_path() -> Path:
    return data_dir() / "bridge_store.json"


def save_mode() -> str:
    return clean_text(os.environ.get("CADASTRO_SAVE_MODE")).lower() or "local"


def save_via_bridge() -> bool:
    return save_mode() in {"bridge", "ponte", "online"}


def bridge_token() -> str:
    return clean_text(os.environ.get("CADASTRO_BRIDGE_TOKEN"))


def token_configured() -> bool:
    return bool(bridge_token())


def verify_token(authorization: str = "") -> bool:
    expected = bridge_token()
    if not expected:
        return False
    authorization = clean_text(authorization)
    prefix = "Bearer "
    token = authorization[len(prefix) :].strip() if authorization.startswith(prefix) else authorization
    return token == expected


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _empty_store() -> dict[str, Any]:
    return {
        "version": BRIDGE_STORE_VERSION,
        "jobs": [],
        "bridge": {
            "last_seen": "",
            "machine": "",
            "workbook_path": "",
            "message": "",
        },
        "products": [],
        "products_updated_at": "",
        "drafts": [],
    }


def _read_store_unlocked() -> dict[str, Any]:
    path = store_path()
    if not path.exists():
        return _empty_store()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        data = _empty_store()
    base = _empty_store()
    base.update(data if isinstance(data, dict) else {})
    base.setdefault("jobs", [])
    base.setdefault("bridge", _empty_store()["bridge"])
    base.setdefault("products", [])
    base.setdefault("drafts", [])
    return base


def _write_store_unlocked(data: dict[str, Any]) -> None:
    path = store_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def read_store() -> dict[str, Any]:
    with _LOCK:
        return deepcopy(_read_store_unlocked())


def form_pairs_from_form(form_data: Any) -> list[dict[str, str]]:
    pairs: list[dict[str, str]] = []
    if hasattr(form_data, "multi_items"):
        iterator = form_data.multi_items()
    elif isinstance(form_data, dict):
        items: list[tuple[Any, Any]] = []
        for key, value in form_data.items():
            if isinstance(value, list):
                items.extend((key, item) for item in value)
            else:
                items.append((key, value))
        iterator = items
    else:
        iterator = []
    for key, value in iterator:
        pairs.append({"name": clean_text(key), "value": clean_text(value)})
    return pairs


class FormPayload:
    def __init__(self, pairs: list[dict[str, str]]):
        self.pairs = pairs

    def get(self, key: str, default: Any = None) -> Any:
        key = clean_text(key)
        for pair in self.pairs:
            if pair.get("name") == key:
                return pair.get("value", "")
        return default

    def getlist(self, key: str) -> list[str]:
        key = clean_text(key)
        return [pair.get("value", "") for pair in self.pairs if pair.get("name") == key]

    def multi_items(self):
        for pair in self.pairs:
            yield pair.get("name", ""), pair.get("value", "")


def enqueue_registration(form_data: Any, category_key: str = "") -> dict[str, Any]:
    pairs = form_pairs_from_form(form_data)
    job = {
        "id": uuid.uuid4().hex[:12],
        "type": "save_registration",
        "status": "queued",
        "created_at": now_text(),
        "updated_at": now_text(),
        "attempts": 0,
        "category_key": clean_text(category_key),
        "payload": {"form_pairs": pairs},
        "result": None,
        "error": "",
    }
    with _LOCK:
        data = _read_store_unlocked()
        data["jobs"].append(job)
        _write_store_unlocked(data)
    return deepcopy(job)


def save_draft(
    category_key: str,
    category_label: str,
    sheet: str,
    descricao_primaria: str,
    payload: dict[str, Any],
    draft_id: str = "",
) -> dict[str, Any]:
    draft_id = clean_text(draft_id) or uuid.uuid4().hex[:12]
    draft = {
        "draft_id": draft_id,
        "category_key": clean_text(category_key),
        "category_label": clean_text(category_label),
        "sheet": clean_text(sheet),
        "saved_at": now_text(),
        "descricao_primaria": clean_text(descricao_primaria) or "(sem descrição primária)",
        "payload": payload or {},
    }
    with _LOCK:
        data = _read_store_unlocked()
        drafts = data.setdefault("drafts", [])
        for index, existing in enumerate(drafts):
            if existing.get("draft_id") == draft_id:
                drafts[index] = draft
                break
        else:
            drafts.append(draft)
        _write_store_unlocked(data)
    return deepcopy(draft)


def list_drafts() -> list[dict[str, Any]]:
    drafts = read_store().get("drafts") or []
    visible = []
    for draft in drafts:
        item = deepcopy(draft)
        item.pop("payload", None)
        visible.append(item)
    return sorted(visible, key=lambda item: item.get("saved_at", ""), reverse=True)


def get_draft(draft_id: str) -> dict[str, Any] | None:
    draft_id = clean_text(draft_id)
    for draft in read_store().get("drafts") or []:
        if draft.get("draft_id") != draft_id:
            continue
        payload = draft.get("payload") or {}
        groups = payload.get("groups") if isinstance(payload, dict) else {}
        return {
            "draft_id": draft.get("draft_id", ""),
            "category_key": draft.get("category_key", ""),
            "category_label": draft.get("category_label", ""),
            "saved_at": draft.get("saved_at", ""),
            "descricao_primaria": draft.get("descricao_primaria", ""),
            "groups": groups if isinstance(groups, dict) else {},
        }
    return None


def delete_draft(draft_id: str) -> dict[str, Any]:
    draft_id = clean_text(draft_id)
    if not draft_id:
        raise ValueError("Rascunho não informado.")
    with _LOCK:
        data = _read_store_unlocked()
        drafts = data.setdefault("drafts", [])
        for index, draft in enumerate(drafts):
            if draft.get("draft_id") != draft_id:
                continue
            removed = drafts.pop(index)
            _write_store_unlocked(data)
            removed.pop("payload", None)
            return deepcopy(removed)
    raise ValueError("Rascunho não encontrado.")


def next_job() -> dict[str, Any] | None:
    with _LOCK:
        data = _read_store_unlocked()
        for job in data["jobs"]:
            if job.get("status") not in {"queued", "retry"}:
                continue
            job["status"] = "processing"
            job["attempts"] = int(job.get("attempts") or 0) + 1
            job["updated_at"] = now_text()
            _write_store_unlocked(data)
            return deepcopy(job)
    return None


def complete_job(job_id: str, result: dict[str, Any]) -> dict[str, Any]:
    with _LOCK:
        data = _read_store_unlocked()
        for job in data["jobs"]:
            if job.get("id") != job_id:
                continue
            job["status"] = "done"
            job["updated_at"] = now_text()
            job["result"] = result
            job["error"] = ""
            _write_store_unlocked(data)
            return deepcopy(job)
    raise ValueError("Job não encontrado.")


def fail_job(job_id: str, error: str, retry: bool = True) -> dict[str, Any]:
    with _LOCK:
        data = _read_store_unlocked()
        for job in data["jobs"]:
            if job.get("id") != job_id:
                continue
            attempts = int(job.get("attempts") or 0)
            job["status"] = "retry" if retry and attempts < 5 else "failed"
            job["updated_at"] = now_text()
            job["error"] = clean_text(error)
            _write_store_unlocked(data)
            return deepcopy(job)
    raise ValueError("Job não encontrado.")


def heartbeat(info: dict[str, Any]) -> dict[str, Any]:
    with _LOCK:
        data = _read_store_unlocked()
        bridge = data.setdefault("bridge", {})
        bridge["last_seen"] = now_text()
        bridge["machine"] = clean_text(info.get("machine"))
        bridge["workbook_path"] = clean_text(info.get("workbook_path"))
        bridge["message"] = clean_text(info.get("message"))
        _write_store_unlocked(data)
        return deepcopy(bridge)


def replace_products(products: list[dict[str, Any]]) -> dict[str, Any]:
    normalized: list[dict[str, str]] = []
    for product in products:
        normalized.append(
            {
                "codigo": clean_text(product.get("codigo")),
                "descricao": clean_text(product.get("descricao")),
                "descricao_secundaria": clean_text(product.get("descricao_secundaria")),
                "unidade": clean_text(product.get("unidade")) or "pc",
                "categoria": clean_text(product.get("categoria")),
            }
        )
    with _LOCK:
        data = _read_store_unlocked()
        data["products"] = normalized
        data["products_updated_at"] = now_text()
        _write_store_unlocked(data)
    return {"count": len(normalized), "updated_at": now_text()}


def products() -> list[dict[str, str]]:
    return read_store().get("products") or []


def status(limit: int = 20) -> dict[str, Any]:
    data = read_store()
    jobs = data.get("jobs") or []
    counts: dict[str, int] = {}
    for job in jobs:
        key = clean_text(job.get("status")) or "unknown"
        counts[key] = counts.get(key, 0) + 1
    return {
        "mode": save_mode(),
        "token_configured": token_configured(),
        "bridge": data.get("bridge") or {},
        "counts": counts,
        "products_count": len(data.get("products") or []),
        "products_updated_at": data.get("products_updated_at") or "",
        "jobs": list(reversed(jobs))[:limit],
        "store_path": str(store_path()),
    }
