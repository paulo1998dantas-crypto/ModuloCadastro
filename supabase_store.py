import json
import os
import re
import tempfile
import urllib.error
import urllib.parse
import urllib.request
import uuid
from io import BytesIO
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

import excel_bancos


REGISTRATIONS_TABLE = "cadastro_registros"
DRAFTS_TABLE = "cadastro_rascunhos"
BOM_HEADERS_TABLE = "cadastro_bom_cabecalhos"
BOM_COMPONENTS_TABLE = "cadastro_bom_componentes"
EXPORT_DIR = Path(tempfile.gettempdir()) / "modulo-cadastro-exports"


class SupabaseStoreError(RuntimeError):
    pass


def clean_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def save_mode() -> str:
    return clean_text(os.environ.get("CADASTRO_SAVE_MODE")).lower()


def enabled() -> bool:
    return save_mode() in {"supabase", "postgres", "database", "banco"}


def configured() -> bool:
    return bool(_supabase_url() and _service_key())


def status() -> dict[str, Any]:
    return {
        "enabled": enabled(),
        "configured": configured(),
        "url": _supabase_url(),
        "tables": [REGISTRATIONS_TABLE, DRAFTS_TABLE, BOM_HEADERS_TABLE, BOM_COMPONENTS_TABLE],
    }


def display_target() -> str:
    if not enabled():
        return ""
    if _supabase_url():
        return f"Modo Supabase: {_supabase_url()} ({REGISTRATIONS_TABLE})"
    return "Modo Supabase: configure SUPABASE_URL e SUPABASE_SERVICE_ROLE_KEY"


def _supabase_url() -> str:
    raw = clean_text(os.environ.get("SUPABASE_URL")) or clean_text(os.environ.get("CADASTRO_SUPABASE_URL"))
    return raw.rstrip("/")


def _service_key() -> str:
    return (
        clean_text(os.environ.get("SUPABASE_SERVICE_ROLE_KEY"))
        or clean_text(os.environ.get("CADASTRO_SUPABASE_SERVICE_ROLE_KEY"))
    )


def _ensure_configured() -> None:
    if not enabled():
        raise SupabaseStoreError("Modo Supabase não está ativo.")
    if not _supabase_url():
        raise SupabaseStoreError("Configure SUPABASE_URL no Render.")
    if not _service_key():
        raise SupabaseStoreError("Configure SUPABASE_SERVICE_ROLE_KEY no Render.")


def _headers(prefer: str = "") -> dict[str, str]:
    key = _service_key()
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
    }
    if prefer:
        headers["Prefer"] = prefer
    return headers


def _request(
    method: str,
    table: str,
    query: list[tuple[str, str]] | None = None,
    payload: Any = None,
    prefer: str = "",
) -> Any:
    _ensure_configured()
    query_string = urllib.parse.urlencode(query or [])
    url = f"{_supabase_url()}/rest/v1/{table}"
    if query_string:
        url = f"{url}?{query_string}"
    data = None
    if payload is not None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(url, data=data, headers=_headers(prefer), method=method)
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            body = response.read().decode("utf-8")
            return json.loads(body) if body else None
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise SupabaseStoreError(f"Erro Supabase {exc.code}: {body}") from exc
    except urllib.error.URLError as exc:
        raise SupabaseStoreError(f"Não foi possível conectar ao Supabase: {exc}") from exc


def _category(category_key: str) -> dict[str, Any]:
    catalog = excel_bancos.load_catalog()
    return excel_bancos._find_category(catalog, category_key)


def _sheet_name(category: dict[str, Any]) -> str:
    return excel_bancos._safe_sheet_title(category.get("sheet_name") or category.get("label") or "Categoria")


def _initial_sku(code_prefix: str) -> str:
    return f"{code_prefix}0001"


def _next_sku(category: dict[str, Any], fields: list[dict[str, Any]], form_data: Any) -> str:
    code_prefix = excel_bancos.pn_code_prefix(category, fields, form_data)
    rows = _request(
        "GET",
        REGISTRATIONS_TABLE,
        [
            ("select", "sku"),
            ("category_key", f"eq.{category['key']}"),
            ("sku", f"like.{code_prefix}%"),
            ("order", "sku.desc"),
            ("limit", "1"),
        ],
    )
    last_code = ""
    if rows:
        last_code = clean_text(rows[0].get("sku"))
    if not last_code or not last_code.isdigit():
        return _initial_sku(code_prefix)
    return str(int(last_code) + 1).zfill(len(last_code))


def _field_groups(fields: list[dict[str, Any]], form_data: Any) -> dict[str, list[str]]:
    groups: dict[str, list[str]] = {}
    for field in fields:
        values = excel_bancos._serialize_field_values(field, form_data)
        groups[field["key"]] = values
    return groups


def _field_values(fields: list[dict[str, Any]], groups: dict[str, list[str]]) -> dict[str, str]:
    values: dict[str, str] = {}
    for field in fields:
        selected = groups.get(field["key"]) or []
        values[field["key"]] = excel_bancos._format_field_saved_value(field, selected) if selected else ""
    return values


def _field_codes(fields: list[dict[str, Any]], groups: dict[str, list[str]]) -> dict[str, str]:
    codes: dict[str, str] = {}
    for field in fields:
        selected = groups.get(field["key"]) or []
        field_codes = [excel_bancos.option_code(value) for value in selected if excel_bancos.option_code(value)]
        codes[field["key"]] = " | ".join(field_codes)
    return codes


def _search_text(*parts: Any) -> str:
    return excel_bancos.normalize_label(" ".join(clean_text(part) for part in parts if clean_text(part)))


def _full_description(row: dict[str, Any]) -> str:
    return " ".join(
        part
        for part in [
            clean_text(row.get("descricao_primaria")),
            clean_text(row.get("descricao_secundaria")),
            clean_text(row.get("sufixo")),
        ]
        if part
    )


def _duplicate_exists(category_key: str, primaria: str, secundaria: str) -> bool:
    rows = _request(
        "GET",
        REGISTRATIONS_TABLE,
        [
            ("select", "id"),
            ("category_key", f"eq.{category_key}"),
            ("descricao_primaria", f"eq.{primaria}"),
            ("descricao_secundaria", f"eq.{secundaria}"),
            ("limit", "1"),
        ],
    )
    return bool(rows)


def save_registration(form_data: Any) -> dict[str, Any]:
    category_key = clean_text(form_data.get("categoria"))
    category = _category(category_key)
    fields = excel_bancos.get_banco_fields(category["key"])
    if category["key"] == excel_bancos.DEFAULT_CATEGORY_KEY:
        excel_bancos._validate_banco_dependencies(fields, form_data)
        excel_bancos._validate_visible_field_requirements(fields, category["key"], form_data)

    descriptions = excel_bancos.build_descriptions(fields, form_data, category["key"])
    if _duplicate_exists(category["key"], descriptions["primaria"], descriptions["secundaria"]):
        raise SupabaseStoreError("Cadastro já existe com a mesma descrição primária e secundária.")

    groups = _field_groups(fields, form_data)
    field_values = _field_values(fields, groups)
    field_codes = _field_codes(fields, groups)
    sku = _next_sku(category, fields, form_data)
    payload = {
        "category_key": category["key"],
        "category_label": category["label"],
        "sheet": _sheet_name(category),
        "sku": sku,
        "descricao_primaria": descriptions["primaria"],
        "descricao_secundaria": descriptions["secundaria"],
        "sufixo": descriptions.get("sufixo") or "",
        "caracteres_primario": len(descriptions["primaria"]),
        "caracteres_secundario": len(descriptions["secundaria"]),
        "form_values": groups,
        "field_values": field_values,
        "field_codes": field_codes,
        "search_text": _search_text(
            sku,
            category["label"],
            descriptions["primaria"],
            descriptions["secundaria"],
            " ".join(field_values.values()),
        ),
    }
    rows = _request("POST", REGISTRATIONS_TABLE, payload=payload, prefer="return=representation")
    row = rows[0] if rows else payload
    return {
        "id": row.get("id"),
        "row": row.get("id") or "-",
        "category": category["label"],
        "category_key": category["key"],
        "sheet": _sheet_name(category),
        "descricao_primaria": descriptions["primaria"],
        "descricao_secundaria": descriptions["secundaria"],
        "sku": sku,
        "path": display_target(),
    }


def _draft_groups(payload_json: str | dict[str, Any]) -> dict[str, list[str]]:
    return excel_bancos._draft_payload_groups(payload_json)


def save_draft(category_key: str, draft_payload: str, draft_id: str = "") -> dict[str, Any]:
    category = _category(category_key)
    fields = excel_bancos.get_banco_fields(category["key"])
    groups = _draft_groups(draft_payload)
    if not groups:
        raise SupabaseStoreError("Rascunho vazio.")
    descriptions = excel_bancos.build_descriptions(fields, groups, category["key"])
    draft_id = clean_text(draft_id) or uuid.uuid4().hex[:12]
    payload = {
        "draft_id": draft_id,
        "category_key": category["key"],
        "category_label": category["label"],
        "sheet": _sheet_name(category),
        "descricao_primaria": descriptions.get("primaria") or "(sem descrição primária)",
        "payload": {"category": category["key"], "groups": groups},
    }
    existing = _request("GET", DRAFTS_TABLE, [("select", "draft_id"), ("draft_id", f"eq.{draft_id}"), ("limit", "1")])
    if existing:
        rows = _request(
            "PATCH",
            DRAFTS_TABLE,
            [("draft_id", f"eq.{draft_id}")],
            payload=payload,
            prefer="return=representation",
        )
    else:
        rows = _request("POST", DRAFTS_TABLE, payload=payload, prefer="return=representation")
    row = rows[0] if rows else payload
    return _draft_summary(row)


def _draft_summary(row: dict[str, Any]) -> dict[str, Any]:
    return {
        "draft_id": clean_text(row.get("draft_id")),
        "category_key": clean_text(row.get("category_key")),
        "category_label": clean_text(row.get("category_label")),
        "sheet": clean_text(row.get("sheet")),
        "saved_at": clean_text(row.get("updated_at") or row.get("created_at")),
        "descricao_primaria": clean_text(row.get("descricao_primaria")),
    }


def list_drafts() -> list[dict[str, Any]]:
    rows = _request("GET", DRAFTS_TABLE, [("select", "*"), ("order", "updated_at.desc")]) or []
    return [_draft_summary(row) for row in rows]


def get_draft(draft_id: str) -> dict[str, Any] | None:
    rows = _request("GET", DRAFTS_TABLE, [("select", "*"), ("draft_id", f"eq.{clean_text(draft_id)}"), ("limit", "1")])
    if not rows:
        return None
    row = rows[0]
    payload = row.get("payload") if isinstance(row.get("payload"), dict) else {}
    return {
        **_draft_summary(row),
        "groups": _draft_groups(payload),
    }


def delete_draft(draft_id: str) -> dict[str, Any]:
    existing = get_draft(draft_id)
    if not existing:
        raise SupabaseStoreError("Rascunho não encontrado.")
    _request("DELETE", DRAFTS_TABLE, [("draft_id", f"eq.{clean_text(draft_id)}")])
    return existing


def _safe_filter_value(value: str) -> str:
    return clean_text(value).replace("*", "").replace(",", " ")


def list_registrations(
    category_key: str = "",
    query: str = "",
    filters: dict[str, str] | None = None,
    limit: int = 250,
    offset: int = 0,
) -> list[dict[str, Any]]:
    selected = _category(category_key)["key"] if clean_text(category_key) else excel_bancos.selected_category("")["key"]
    params: list[tuple[str, str]] = [
        ("select", "*"),
        ("category_key", f"eq.{selected}"),
        ("order", "sku.asc"),
        ("limit", str(max(1, min(limit, 2000)))),
        ("offset", str(max(0, offset))),
    ]
    term = _search_text(query)
    if term:
        params.append(("search_text", f"ilike.*{term}*"))
    for key, value in (filters or {}).items():
        value = _safe_filter_value(value)
        if key and value:
            params.append((f"field_values->>{key}", f"ilike.*{value}*"))
    return _request("GET", REGISTRATIONS_TABLE, params) or []


def get_registration(registration_id: int | str) -> dict[str, Any] | None:
    rows = _request(
        "GET",
        REGISTRATIONS_TABLE,
        [
            ("select", "*"),
            ("id", f"eq.{clean_text(registration_id)}"),
            ("limit", "1"),
        ],
    )
    return rows[0] if rows else None


def _groups_from_record(fields: list[dict[str, Any]], record: dict[str, Any]) -> dict[str, list[str]]:
    form_values = record.get("form_values") if isinstance(record.get("form_values"), dict) else {}
    field_values = record.get("field_values") if isinstance(record.get("field_values"), dict) else {}
    groups: dict[str, list[str]] = {}
    for field in fields:
        raw = form_values.get(field["key"])
        if isinstance(raw, list) and raw:
            groups[field["key"]] = [clean_text(value) for value in raw if clean_text(value)]
            continue
        saved = clean_text(field_values.get(field["key"]))
        if saved:
            groups[field["key"]] = [value.strip() for value in saved.split("|") if value.strip()]
    return groups


def editable_registration(registration_id: int | str) -> dict[str, Any]:
    record = get_registration(registration_id)
    if not record:
        raise SupabaseStoreError("Cadastro não encontrado.")
    category = _category(clean_text(record.get("category_key")))
    fields = excel_bancos.get_banco_fields(category["key"])
    groups = _groups_from_record(fields, record)
    return {"record": record, "category": category, "fields": fields, "groups": groups}


def update_registration(registration_id: int | str, form_data: Any) -> dict[str, Any]:
    current = get_registration(registration_id)
    if not current:
        raise SupabaseStoreError("Cadastro não encontrado.")
    category = _category(clean_text(current.get("category_key")))
    fields = excel_bancos.get_banco_fields(category["key"])
    if category["key"] == excel_bancos.DEFAULT_CATEGORY_KEY:
        excel_bancos._validate_banco_dependencies(fields, form_data)
        excel_bancos._validate_visible_field_requirements(fields, category["key"], form_data)

    descriptions = excel_bancos.build_descriptions(fields, form_data, category["key"])
    groups = _field_groups(fields, form_data)
    field_values = _field_values(fields, groups)
    field_codes = _field_codes(fields, groups)
    sku = clean_text(current.get("sku"))
    payload = {
        "category_label": category["label"],
        "sheet": _sheet_name(category),
        "descricao_primaria": descriptions["primaria"],
        "descricao_secundaria": descriptions["secundaria"],
        "sufixo": descriptions.get("sufixo") or "",
        "caracteres_primario": len(descriptions["primaria"]),
        "caracteres_secundario": len(descriptions["secundaria"]),
        "form_values": groups,
        "field_values": field_values,
        "field_codes": field_codes,
        "search_text": _search_text(
            sku,
            category["label"],
            descriptions["primaria"],
            descriptions["secundaria"],
            " ".join(field_values.values()),
        ),
    }
    rows = _request(
        "PATCH",
        REGISTRATIONS_TABLE,
        [("id", f"eq.{clean_text(registration_id)}")],
        payload=payload,
        prefer="return=representation",
    )
    return rows[0] if rows else {**current, **payload}


def search_products(query: str, limit: int = 25) -> list[dict[str, str]]:
    term = _search_text(query)
    if len(term) < 1:
        return []
    rows = _request(
        "GET",
        REGISTRATIONS_TABLE,
        [
            ("select", "sku,descricao_primaria,category_label,search_text"),
            ("search_text", f"ilike.*{term}*"),
            ("order", "sku.asc"),
            ("limit", str(limit)),
        ],
    ) or []
    return [
        {
            "codigo": clean_text(row.get("sku")),
            "descricao": clean_text(row.get("descricao_primaria")),
            "categoria": clean_text(row.get("category_label")),
            "unidade": "pc",
        }
        for row in rows
    ]


def _registration_by_sku(sku: str) -> dict[str, Any] | None:
    rows = _request(
        "GET",
        REGISTRATIONS_TABLE,
        [
            ("select", "id,category_key,category_label,sku,descricao_primaria,descricao_secundaria,sufixo"),
            ("sku", f"eq.{clean_text(sku)}"),
            ("limit", "1"),
        ],
    )
    return rows[0] if rows else None


def _bom_header_by_parent(parent_sku: str) -> dict[str, Any] | None:
    rows = _request(
        "GET",
        BOM_HEADERS_TABLE,
        [
            ("select", "*"),
            ("parent_sku", f"eq.{clean_text(parent_sku)}"),
            ("limit", "1"),
        ],
    )
    return rows[0] if rows else None


def save_bom(
    parent_sku: str,
    parent_description: str,
    components: list[dict[str, Any]],
    category_key: str = "",
    category_label: str = "",
    registration_id: int | str | None = None,
    source: str = "cadastro",
) -> dict[str, Any]:
    parent_sku = clean_text(parent_sku)
    if not parent_sku:
        raise SupabaseStoreError("Informe o SKU do item pai da B.O.M.")
    if not components:
        raise SupabaseStoreError("Informe pelo menos um componente para a B.O.M.")

    registration = _registration_by_sku(parent_sku)
    if registration:
        category_key = clean_text(registration.get("category_key")) or category_key
        category_label = clean_text(registration.get("category_label")) or category_label
        parent_description = _full_description(registration) or parent_description
        registration_id = registration.get("id") or registration_id

    parent_description = clean_text(parent_description) or parent_sku
    payload = {
        "parent_sku": parent_sku,
        "parent_descricao": parent_description,
        "parent_category_key": clean_text(category_key),
        "parent_category_label": clean_text(category_label),
        "registration_id": registration_id,
        "source": clean_text(source) or "cadastro",
        "search_text": _search_text(parent_sku, parent_description, category_label),
    }

    existing = _bom_header_by_parent(parent_sku)
    if existing:
        rows = _request(
            "PATCH",
            BOM_HEADERS_TABLE,
            [("id", f"eq.{existing['id']}")],
            payload=payload,
            prefer="return=representation",
        )
    else:
        rows = _request("POST", BOM_HEADERS_TABLE, payload=payload, prefer="return=representation")
    header = rows[0] if rows else {**payload, "id": existing.get("id") if existing else None}
    bom_id = header.get("id")
    if not bom_id:
        raise SupabaseStoreError("Nao foi possivel criar o cabecalho da B.O.M.")

    _request("DELETE", BOM_COMPONENTS_TABLE, [("bom_id", f"eq.{bom_id}")])
    component_payloads = []
    for index, component in enumerate(components, start=1):
        component_sku = clean_text(component.get("codigo") or component.get("component_sku"))
        if not component_sku:
            continue
        try:
            quantity = float(component.get("quantidade") or component.get("quantity") or 0)
        except Exception as exc:
            raise SupabaseStoreError(f"Quantidade invalida no componente {component_sku}.") from exc
        if quantity <= 0:
            raise SupabaseStoreError(f"Quantidade deve ser maior que zero no componente {component_sku}.")
        description = clean_text(component.get("descricao") or component.get("component_descricao"))
        unit = clean_text(component.get("unidade") or component.get("unit")) or "pc"
        component_payloads.append(
            {
                "bom_id": bom_id,
                "parent_sku": parent_sku,
                "component_sku": component_sku,
                "component_descricao": description,
                "unidade": unit,
                "quantidade": quantity,
                "ordem": index,
                "search_text": _search_text(parent_sku, parent_description, component_sku, description, unit),
            }
        )
    if not component_payloads:
        raise SupabaseStoreError("Informe pelo menos um componente valido para a B.O.M.")
    _request("POST", BOM_COMPONENTS_TABLE, payload=component_payloads, prefer="return=minimal")
    return {"bom": header, "components_count": len(component_payloads)}


def _in_filter(values: list[Any]) -> str:
    cleaned = [clean_text(value) for value in values if clean_text(value)]
    return "in.(" + ",".join(cleaned) + ")"


def list_boms(
    category_key: str = "",
    parent_query: str = "",
    component_query: str = "",
    limit: int = 500,
) -> list[dict[str, Any]]:
    bom_ids_filter: list[str] = []
    component_term = _search_text(component_query)
    if component_term:
        component_matches = _request(
            "GET",
            BOM_COMPONENTS_TABLE,
            [
                ("select", "bom_id"),
                ("search_text", f"ilike.*{component_term}*"),
                ("limit", "5000"),
            ],
        ) or []
        bom_ids_filter = list(dict.fromkeys(clean_text(row.get("bom_id")) for row in component_matches if row.get("bom_id")))
        if not bom_ids_filter:
            return []

    params: list[tuple[str, str]] = [
        ("select", "*"),
        ("order", "parent_sku.asc"),
        ("limit", str(max(1, min(limit, 5000)))),
    ]
    if clean_text(category_key):
        params.append(("parent_category_key", f"eq.{clean_text(category_key)}"))
    parent_term = _search_text(parent_query)
    if parent_term:
        params.append(("search_text", f"ilike.*{parent_term}*"))
    if bom_ids_filter:
        params.append(("id", _in_filter(bom_ids_filter)))
    headers = _request("GET", BOM_HEADERS_TABLE, params) or []
    bom_ids = [clean_text(row.get("id")) for row in headers if row.get("id")]
    components_by_bom: dict[str, list[dict[str, Any]]] = {bom_id: [] for bom_id in bom_ids}
    if bom_ids:
        component_params: list[tuple[str, str]] = [
            ("select", "*"),
            ("bom_id", _in_filter(bom_ids)),
            ("order", "parent_sku.asc,ordem.asc,component_sku.asc"),
            ("limit", "10000"),
        ]
        if component_term:
            component_params.append(("search_text", f"ilike.*{component_term}*"))
        components = _request("GET", BOM_COMPONENTS_TABLE, component_params) or []
        for component in components:
            components_by_bom.setdefault(clean_text(component.get("bom_id")), []).append(component)
    return [{**header, "components": components_by_bom.get(clean_text(header.get("id")), [])} for header in headers]


def get_bom(bom_id: int | str) -> dict[str, Any]:
    bom_id = clean_text(bom_id)
    rows = _request("GET", BOM_HEADERS_TABLE, [("select", "*"), ("id", f"eq.{bom_id}"), ("limit", "1")]) or []
    if not rows:
        raise SupabaseStoreError("B.O.M. nao encontrada.")
    components = _request(
        "GET",
        BOM_COMPONENTS_TABLE,
        [
            ("select", "*"),
            ("bom_id", f"eq.{bom_id}"),
            ("order", "ordem.asc,component_sku.asc"),
            ("limit", "10000"),
        ],
    ) or []
    return {**rows[0], "components": components}


def update_bom(bom_id: int | str, parent_description: str, components: list[dict[str, Any]]) -> dict[str, Any]:
    current = get_bom(bom_id)
    parent_sku = clean_text(current.get("parent_sku"))
    parent_description = clean_text(parent_description) or clean_text(current.get("parent_descricao")) or parent_sku
    if not components:
        raise SupabaseStoreError("Informe pelo menos um componente para a B.O.M.")
    header_payload = {
        "parent_descricao": parent_description,
        "source": "edicao",
        "search_text": _search_text(parent_sku, parent_description, current.get("parent_category_label")),
    }
    rows = _request(
        "PATCH",
        BOM_HEADERS_TABLE,
        [("id", f"eq.{clean_text(bom_id)}")],
        payload=header_payload,
        prefer="return=representation",
    )
    header = rows[0] if rows else {**current, **header_payload}

    _request("DELETE", BOM_COMPONENTS_TABLE, [("bom_id", f"eq.{clean_text(bom_id)}")])
    component_payloads = []
    for index, component in enumerate(components, start=1):
        component_sku = clean_text(component.get("codigo") or component.get("component_sku"))
        if not component_sku:
            continue
        try:
            quantity = float(component.get("quantidade") or component.get("quantity") or 0)
        except Exception as exc:
            raise SupabaseStoreError(f"Quantidade invalida no componente {component_sku}.") from exc
        if quantity <= 0:
            raise SupabaseStoreError(f"Quantidade deve ser maior que zero no componente {component_sku}.")
        description = clean_text(component.get("descricao") or component.get("component_descricao"))
        unit = clean_text(component.get("unidade") or component.get("unit")) or "pc"
        component_payloads.append(
            {
                "bom_id": clean_text(bom_id),
                "parent_sku": parent_sku,
                "component_sku": component_sku,
                "component_descricao": description,
                "unidade": unit,
                "quantidade": quantity,
                "ordem": index,
                "search_text": _search_text(parent_sku, parent_description, component_sku, description, unit),
            }
        )
    if not component_payloads:
        raise SupabaseStoreError("Informe pelo menos um componente valido para a B.O.M.")
    _request("POST", BOM_COMPONENTS_TABLE, payload=component_payloads, prefer="return=minimal")
    return {**header, "components": component_payloads}


def delete_bom(bom_id: int | str) -> dict[str, Any]:
    bom_id = clean_text(bom_id)
    rows = _request("GET", BOM_HEADERS_TABLE, [("select", "*"), ("id", f"eq.{bom_id}"), ("limit", "1")]) or []
    if not rows:
        raise SupabaseStoreError("B.O.M. nao encontrada.")
    _request("DELETE", BOM_COMPONENTS_TABLE, [("bom_id", f"eq.{bom_id}")])
    _request("DELETE", BOM_HEADERS_TABLE, [("id", f"eq.{bom_id}")])
    return rows[0]


def _normalize_header(value: Any) -> str:
    return excel_bancos.normalize_label(value).replace(" ", "_").lower()


def import_bom_workbook(content: bytes, filename: str = "") -> dict[str, Any]:
    wb = load_workbook(BytesIO(content), data_only=True)
    imported = 0
    parents: dict[str, dict[str, Any]] = {}
    try:
        for ws in wb.worksheets:
            header_row = None
            header_map: dict[str, int] = {}
            for row_index in range(1, min(ws.max_row, 10) + 1):
                headers = {_normalize_header(ws.cell(row_index, col).value): col for col in range(1, ws.max_column + 1)}
                if "item_codigo" in headers and "componente_codigo" in headers:
                    header_row = row_index
                    header_map = headers
                    break
            if not header_row:
                continue
            item_block_col = header_map.get("item_bloco")
            parent_hint = clean_text(ws.cell(header_row, item_block_col + 1).value) if item_block_col else ""
            for row_index in range(header_row + 1, ws.max_row + 1):
                parent_sku = clean_text(ws.cell(row_index, header_map["item_codigo"]).value)
                component_sku = clean_text(ws.cell(row_index, header_map["componente_codigo"]).value)
                if not parent_sku or not component_sku:
                    continue
                description = clean_text(ws.cell(row_index, header_map.get("descricao", 3)).value)
                unit = clean_text(ws.cell(row_index, header_map.get("unidade", 4)).value) or "pc"
                quantity = ws.cell(row_index, header_map.get("quantidade", 5)).value
                group = parents.setdefault(
                    parent_sku,
                    {
                        "parent_description": parent_hint or parent_sku,
                        "components": [],
                    },
                )
                group["components"].append(
                    {
                        "codigo": component_sku,
                        "descricao": description,
                        "unidade": unit,
                        "quantidade": quantity,
                    }
                )
        for parent_sku, data in parents.items():
            save_bom(parent_sku, data["parent_description"], data["components"], source=f"import:{filename or 'xlsx'}")
            imported += 1
    finally:
        wb.close()
    return {"parents": imported, "components": sum(len(item["components"]) for item in parents.values())}


def export_boms(category_key: str = "", parent_query: str = "", component_query: str = "") -> Path:
    rows = list_boms(category_key=category_key, parent_query=parent_query, component_query=component_query, limit=5000)
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    output = EXPORT_DIR / f"bom_{stamp}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "BOM"
    headers = [
        "categoria",
        "item_codigo",
        "item_descricao",
        "componente_codigo",
        "descricao",
        "unidade",
        "quantidade",
    ]
    ws.append(headers)
    for bom in rows:
        for component in bom.get("components") or []:
            quantity = component.get("quantidade")
            try:
                quantity = int(quantity) if float(quantity).is_integer() else float(quantity)
            except Exception:
                pass
            ws.append(
                [
                    bom.get("parent_category_label"),
                    bom.get("parent_sku"),
                    bom.get("parent_descricao"),
                    component.get("component_sku"),
                    component.get("component_descricao"),
                    component.get("unidade"),
                    quantity,
                ]
            )
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="E2E8F0")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for column, width in {"A": 24, "B": 16, "C": 54, "D": 20, "E": 72, "F": 12, "G": 14}.items():
        ws.column_dimensions[column].width = width
    for row_cells in ws.iter_rows(min_row=2):
        for cell in row_cells:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"
    wb.save(output)
    wb.close()
    return output


def export_registrations(category_key: str, query: str = "", filters: dict[str, str] | None = None) -> Path:
    category = _category(category_key)
    fields = excel_bancos.get_banco_fields_for_display(category["key"])
    rows = list_registrations(category["key"], query=query, filters=filters, limit=10000)
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    output = EXPORT_DIR / f"cadastros_{category['key']}_{stamp}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = _sheet_name(category)[:31]
    headers = [
        "SKU",
        "DESCRIÇÃO PRIMÁRIA",
        "DESCRIÇÃO SECUNDÁRIA",
        "SUFIXO",
        "CARACTERES PRIMARIO",
        "CARACTERES SECUNDARIO",
    ]
    headers.extend(excel_bancos.header_for_field(field["label"], field["scope"]) for field in fields)
    ws.append(headers)
    for row in rows:
        values = row.get("field_values") if isinstance(row.get("field_values"), dict) else {}
        ws.append(
            [
                row.get("sku"),
                row.get("descricao_primaria"),
                row.get("descricao_secundaria"),
                row.get("sufixo"),
                row.get("caracteres_primario"),
                row.get("caracteres_secundario"),
                *[values.get(field["key"], "") for field in fields],
            ]
        )

    header_fill = PatternFill("solid", fgColor="E2E8F0")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    widths = {1: 14, 2: 48, 3: 72, 4: 18, 5: 18, 6: 22}
    for index in range(1, len(headers) + 1):
        ws.column_dimensions[ws.cell(1, index).column_letter].width = widths.get(index, 28)
    for row_cells in ws.iter_rows(min_row=2):
        for cell in row_cells:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"
    wb.save(output)
    wb.close()
    return output
