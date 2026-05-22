import json
import os
import re
import shutil
import sys
import tempfile
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from copy import copy, deepcopy
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.formula import ArrayFormula


DEFAULT_NEW_WORKBOOK_NAME = "Cadastro Bancos.xlsx"
DEFAULT_CATEGORY_KEY = "bancos"
DEFAULT_CATEGORY_LABEL = "Bancos"
FIRST_DATA_ROW = 3
REGISTRATION_SHEET_NAME = "_cadastro_app"
SELECTION_MODE_UNITARIA = "unitaria"
SELECTION_MODE_MULTIPLA = "multipla"
EXCEL_SUFFIXES = {".xlsx", ".xlsm"}
IGNORED_SCAN_DIRS = {"__pycache__", "outputs", ".idea", ".venv", "backups", "browser_profile", "_internal"}
VALUE_RE = re.compile(r"^\s*(\d+)\s*[-–—]\s*(.+?)\s*$")
INVALID_SHEET_CHARS_RE = re.compile(r'[:\\/?*\[\]]')
DISTANCIA_PE_KEY = "distancia_pe"
DISTANCIA_PE_PREFIX = "ORIENTADO A ESQ:"
ORDINAL_MASCULINE = "\ufff0"
ORDINAL_FEMININE = "\ufff1"
DESCRIPTION_PRIMARY_HEADER = "DESCRICAO PRIMARIA"
DESCRIPTION_SECONDARY_HEADER = "DESCRICAO SECUNDARIA"
DESCRIPTION_SUFFIX_HEADER = "SUFIXO"
_CATALOG_CACHE: dict[str, Any] | None = None
_CATALOG_CACHE_MTIME: float | None = None
_REGISTRATION_CACHE: dict[tuple[str, str, int], dict[str, Any]] = {}


def _project_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


PROJECT_DIR = _project_dir()
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", PROJECT_DIR))
CONFIG_PATH = PROJECT_DIR / "config.json"
DATA_PATH = PROJECT_DIR / "cadastro_dados.json"
BACKUP_DIR = PROJECT_DIR / "backups"
_template_env = os.environ.get("CADASTRO_BANCO_TEMPLATE", "").strip()
DEFAULT_TEMPLATE_PATH = Path(_template_env).expanduser() if _template_env else PROJECT_DIR / "TEMPLATE ID BANCO.xlsx"


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_label(value: Any) -> str:
    text = re.sub(r"\([^)]*\)", "", clean_text(value))
    text = unicodedata.normalize("NFKD", text.upper())
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^A-Z0-9]+", " ", text).strip()


def option_label(value: Any) -> str:
    text = clean_text(value)
    match = VALUE_RE.match(text)
    if match:
        return match.group(2).strip()
    return text


def option_code(value: Any) -> str:
    text = clean_text(value)
    match = VALUE_RE.match(text)
    if match:
        return match.group(1).strip()
    return ""


def option_identity(value: Any) -> str:
    return normalize_label(option_label(value))


def strip_accents(value: Any) -> str:
    text = clean_text(value)
    normalized = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in normalized if not unicodedata.combining(ch))


def normalize_option_text(value: Any) -> str:
    text = clean_text(value)
    if not text:
        return ""

    ordinal_masculine = "º"
    ordinal_feminine = "ª"

    def _preserve_ordinals(input_text: str) -> str:
        return input_text.replace(ordinal_masculine, "￰").replace(ordinal_feminine, "￱")

    def _restore_ordinals(input_text: str) -> str:
        return input_text.replace("￰", ordinal_masculine).replace("￱", ordinal_feminine)

    match = VALUE_RE.match(text)
    if match:
        code = match.group(1).strip()
        body = _preserve_ordinals(match.group(2).strip())
        body = strip_accents(body).upper()
        body = _restore_ordinals(body)
        body = re.sub(r"\s+", " ", body)
        return f"{code}- {body}"

    text = _preserve_ordinals(text)
    text = strip_accents(text).upper()
    text = _restore_ordinals(text)
    return re.sub(r"\s+", " ", text).strip()

def compose_secondary_description(primary: Any, secondary: Any) -> str:
    primary_text = clean_text(primary)
    secondary_text = clean_text(secondary)
    if not primary_text:
        return secondary_text
    if not secondary_text:
        return primary_text

    normalized_primary = normalize_label(primary_text)
    normalized_secondary = normalize_label(secondary_text)
    if normalized_secondary == normalized_primary or normalized_secondary.startswith(f"{normalized_primary} "):
        return secondary_text
    return f"{primary_text} {secondary_text}".strip()


def field_key(value: Any) -> str:
    normalized = normalize_label(value).lower().replace(" ", "_")
    return normalized or "campo"


def category_key(value: Any) -> str:
    normalized = normalize_label(value).lower().replace(" ", "_")
    return normalized or "categoria"


def field_scope(value: Any) -> str:
    return "secundaria" if clean_text(value).lower() == "secundaria" else "primaria"


def field_selection_mode(value: Any) -> str:
    normalized = normalize_label(value)
    if normalized in {"MULTIPLA", "MULTIPLA SELECAO", "SELECAO MULTIPLA"}:
        return SELECTION_MODE_MULTIPLA
    return SELECTION_MODE_UNITARIA


def header_for_field(label: str, scope: str) -> str:
    suffix = "PRIMARIO" if field_scope(scope) == "primaria" else "SECUNDARIO"
    return f"{clean_text(label)} - {suffix}"


def _read_json(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError:
        return {}


def _write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _read_config() -> dict[str, Any]:
    return _read_json(CONFIG_PATH)


def _write_config(data: dict[str, Any]) -> None:
    _write_json(CONFIG_PATH, data)


def _seed_candidates(file_name: str) -> list[Path]:
    return [
        PROJECT_DIR / "_internal" / file_name,
        RESOURCE_DIR / file_name,
        RESOURCE_DIR / "_internal" / file_name,
    ]


def _ensure_seed_file(path: Path, file_name: str) -> None:
    if path.exists():
        return
    for candidate in _seed_candidates(file_name):
        if candidate.exists():
            shutil.copy2(candidate, path)
            return


def _copy_to_temp(path: Path) -> Path:
    suffix = path.suffix or ".xlsx"
    fd, temp_name = tempfile.mkstemp(prefix="_cadastro_app_", suffix=suffix)
    os.close(fd)
    temp_path = Path(temp_name)
    shutil.copy2(path, temp_path)
    return temp_path


def _load(path: Path):
    try:
        return load_workbook(path, keep_links=True, keep_vba=False)
    except IndexError:
        repaired_path = _repair_invalid_style_ids(path)
        workbook = load_workbook(repaired_path, keep_links=True, keep_vba=False)
        setattr(workbook, "_cadastro_cleanup_path", repaired_path)
        return workbook


def _repair_invalid_style_ids(path: Path) -> Path:
    if not path.exists():
        return path

    fd, temp_name = tempfile.mkstemp(prefix="_cadastro_app_repair_", suffix=path.suffix or ".xlsx", dir=str(path.parent))
    os.close(fd)
    temp_path = Path(temp_name)
    ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

    try:
        with zipfile.ZipFile(path, "r") as source_zip:
            try:
                styles_root = ET.fromstring(source_zip.read("xl/styles.xml"))
                cellxfs = styles_root.find("a:cellXfs", ns)
                style_count = int(cellxfs.attrib.get("count", "0")) if cellxfs is not None else 0
            except Exception:
                style_count = 0

            max_style = max(0, style_count - 1)

            with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as output_zip:
                for info in source_zip.infolist():
                    payload = source_zip.read(info.filename)
                    if info.filename.startswith("xl/worksheets/") and info.filename.endswith(".xml") and style_count:
                        try:
                            root = ET.fromstring(payload)
                            changed = False
                            for cell in root.findall(".//a:c", ns):
                                style_value = cell.attrib.get("s")
                                if style_value is None:
                                    continue
                                try:
                                    style_id = int(style_value)
                                except ValueError:
                                    continue
                                if style_id >= style_count:
                                    cell.set("s", str(max_style))
                                    changed = True
                            if changed:
                                payload = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                        except Exception:
                            pass
                    output_zip.writestr(info, payload)
        return temp_path
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise


def _unique_key(label: str, used: set[str]) -> str:
    base = field_key(label)
    candidate = base
    counter = 2
    while candidate in used:
        candidate = f"{base}_{counter}"
        counter += 1
    return candidate


def _unique_category_key(label: str, used: set[str]) -> str:
    base = category_key(label)
    candidate = base
    counter = 2
    while candidate in used:
        candidate = f"{base}_{counter}"
        counter += 1
    return candidate


def _default_category() -> dict[str, Any]:
    return {
        "key": DEFAULT_CATEGORY_KEY,
        "label": DEFAULT_CATEGORY_LABEL,
        "sheet_name": DEFAULT_CATEGORY_LABEL,
        "fields": [],
    }


def _default_catalog() -> dict[str, Any]:
    return {
        "version": 2,
        "active_category": DEFAULT_CATEGORY_KEY,
        "categories": [_default_category()],
    }


def _sanitize_fields(fields: list[dict[str, Any]] | None, include_defaults: bool = False) -> list[dict[str, Any]]:
    used_keys: set[str] = set()
    cleaned: list[dict[str, Any]] = []
    for field in fields or []:
        label = clean_text(field.get("label")).upper()
        if not label:
            continue
        key = clean_text(field.get("key")) or field_key(label)
        if key in used_keys:
            key = _unique_key(label, used_keys)
        used_keys.add(key)

        scope = field_scope(field.get("scope"))
        selection_mode = field_selection_mode(field.get("selection_mode"))

        options: list[str] = []
        seen_options: set[str] = set()
        for option in field.get("options") or []:
            value = normalize_option_text(option)
            if not value:
                continue
            option_key = option_identity(value)
            if option_key in seen_options:
                continue
            seen_options.add(option_key)
            options.append(value)

        cleaned.append(
            {
                "key": key,
                "label": label,
                "scope": scope,
                "selection_mode": selection_mode,
                "options": options,
            }
        )

    if include_defaults and not cleaned:
        cleaned.extend(_default_category()["fields"])
    return cleaned


def _sanitize_catalog(catalog: dict[str, Any]) -> dict[str, Any]:
    if catalog.get("fields") is not None:
        raw_categories = [
            {
                "key": DEFAULT_CATEGORY_KEY,
                "label": DEFAULT_CATEGORY_LABEL,
                "fields": catalog.get("fields") or [],
            }
        ]
        active_category_value = DEFAULT_CATEGORY_KEY
    else:
        raw_categories = catalog.get("categories") or []
        active_category_value = clean_text(catalog.get("active_category")) or DEFAULT_CATEGORY_KEY

    if not raw_categories:
        raw_categories = [_default_category()]
        active_category_value = DEFAULT_CATEGORY_KEY

    categories: list[dict[str, Any]] = []
    used_keys: set[str] = set()
    used_labels: set[str] = set()

    for raw_category in raw_categories:
        label = clean_text(raw_category.get("label")) or "Categoria"
        label_key = normalize_label(label)
        if label_key in used_labels:
            label = f"{label} {len(categories) + 1}"
            label_key = normalize_label(label)
        used_labels.add(label_key)

        raw_key = clean_text(raw_category.get("key")) or category_key(label)
        key = raw_key if raw_key not in used_keys else _unique_category_key(label, used_keys)
        used_keys.add(key)

        categories.append(
            {
                "key": key,
                "label": label,
                "sheet_name": _safe_sheet_title(raw_category.get("sheet_name") or label),
                "fields": _sanitize_fields(raw_category.get("fields") or [], include_defaults=False),
            }
        )

    if active_category_value not in {category["key"] for category in categories}:
        active_category_value = categories[0]["key"]

    return {
        "version": 2,
        "active_category": active_category_value,
        "categories": categories,
    }


def load_catalog() -> dict[str, Any]:
    global _CATALOG_CACHE, _CATALOG_CACHE_MTIME
    _ensure_seed_file(DATA_PATH, "cadastro_dados.json")
    if not DATA_PATH.exists():
        catalog = _default_catalog()
        save_catalog(catalog)
        return deepcopy(catalog)

    mtime = DATA_PATH.stat().st_mtime
    if _CATALOG_CACHE is not None and _CATALOG_CACHE_MTIME == mtime:
        return deepcopy(_CATALOG_CACHE)

    raw_catalog = _read_json(DATA_PATH)
    catalog = _sanitize_catalog(raw_catalog)
    if raw_catalog != catalog:
        save_catalog(catalog)
        mtime = DATA_PATH.stat().st_mtime

    _CATALOG_CACHE = catalog
    _CATALOG_CACHE_MTIME = mtime
    return deepcopy(catalog)


def save_catalog(catalog: dict[str, Any]) -> None:
    global _CATALOG_CACHE, _CATALOG_CACHE_MTIME
    sanitized = _sanitize_catalog(catalog)
    _write_json(DATA_PATH, sanitized)
    _CATALOG_CACHE = sanitized
    _CATALOG_CACHE_MTIME = DATA_PATH.stat().st_mtime


def list_categories() -> list[dict[str, Any]]:
    categories = []
    for category in load_catalog()["categories"]:
        categories.append(
            {
                "key": category["key"],
                "label": category["label"],
                "fields_count": len(category.get("fields") or []),
            }
        )
    return categories


def _find_category(catalog: dict[str, Any], category_key_value: str) -> dict[str, Any]:
    requested = clean_text(category_key_value) or clean_text(catalog.get("active_category")) or DEFAULT_CATEGORY_KEY
    for category in catalog["categories"]:
        if category["key"] == requested:
            return category
    return catalog["categories"][0]


def selected_category(category_key_value: str) -> dict[str, Any]:
    category = _find_category(load_catalog(), category_key_value)
    return {"key": category["key"], "label": category["label"]}


def set_active_category(category_key_value: str) -> dict[str, Any]:
    catalog = load_catalog()
    category = _find_category(catalog, category_key_value)
    catalog["active_category"] = category["key"]
    save_catalog(catalog)
    sync_workbook_structure(category["key"])
    return {"key": category["key"], "label": category["label"]}


def add_category(label: str) -> dict[str, str]:
    catalog = load_catalog()
    label_clean = clean_text(label)
    if not label_clean:
        raise ValueError("Informe o nome da categoria.")
    if normalize_label(label_clean) in {normalize_label(item["label"]) for item in catalog["categories"]}:
        raise ValueError("Essa categoria já existe.")
    used = {item["key"] for item in catalog["categories"]}
    category = {
        "key": _unique_category_key(label_clean, used),
        "label": label_clean,
        "sheet_name": _safe_sheet_title(label_clean),
        "fields": [],
    }
    catalog["categories"].append(category)
    catalog["active_category"] = category["key"]
    save_catalog(catalog)
    sync_workbook_structure(category["key"])
    return {
        "category": category["label"],
        "category_key": category["key"],
        "path": str(DATA_PATH),
    }


def update_category(category_key_value: str, label: str) -> dict[str, str]:
    catalog = load_catalog()
    category = _find_category(catalog, category_key_value)
    label_clean = clean_text(label)
    if not label_clean:
        raise ValueError("Informe o nome da categoria.")
    for existing in catalog["categories"]:
        if existing["key"] != category["key"] and normalize_label(existing["label"]) == normalize_label(label_clean):
            raise ValueError("Já existe outra categoria com esse nome.")
    category["label"] = label_clean
    category["sheet_name"] = _safe_sheet_title(label_clean)
    save_catalog(catalog)
    sync_workbook_structure(category["key"])
    return {
        "category": category["label"],
        "category_key": category["key"],
        "path": str(DATA_PATH),
    }


def delete_category(category_key_value: str) -> dict[str, str]:
    catalog = load_catalog()
    category = _find_category(catalog, category_key_value)
    if len(catalog["categories"]) == 1:
        raise ValueError("Pelo menos uma categoria deve permanecer cadastrada.")
    catalog["categories"] = [item for item in catalog["categories"] if item["key"] != category["key"]]
    if catalog.get("active_category") == category["key"]:
        catalog["active_category"] = catalog["categories"][0]["key"]
    save_catalog(catalog)
    return {"category": category["label"], "path": str(DATA_PATH)}


def active_workbook_path() -> Path:
    _ensure_seed_file(CONFIG_PATH, "config.json")
    config = _read_config()
    raw = clean_text(config.get("active_workbook"))
    if raw:
        path = Path(raw)
        if not path.is_absolute():
            path = (PROJECT_DIR / path).resolve()
        return path
    return (PROJECT_DIR / DEFAULT_NEW_WORKBOOK_NAME).resolve()


def template_path() -> str:
    return str(active_workbook_path())


def _iter_project_paths() -> list[Path]:
    found: list[Path] = []
    for root, dirs, files in os.walk(PROJECT_DIR):
        dirs[:] = [name for name in dirs if name not in IGNORED_SCAN_DIRS]
        root_path = Path(root)
        for name in files:
            found.append(root_path / name)
    return found


def list_workbooks() -> list[dict[str, str]]:
    workbooks: list[dict[str, str]] = []
    for path in _iter_project_paths():
        if path.suffix.lower() not in EXCEL_SUFFIXES:
            continue
        workbooks.append(
            {
                "path": str(path),
                "name": path.name,
                "relative": str(path.relative_to(PROJECT_DIR)),
            }
        )
    return sorted(workbooks, key=lambda item: item["relative"].lower())


def list_folders() -> list[dict[str, str]]:
    folders: list[dict[str, str]] = []
    for root, dirs, _ in os.walk(PROJECT_DIR):
        dirs[:] = [name for name in dirs if name not in IGNORED_SCAN_DIRS]
        path = Path(root)
        if path == PROJECT_DIR:
            continue
        folders.append(
            {
                "path": str(path),
                "name": path.name,
                "relative": str(path.relative_to(PROJECT_DIR)),
            }
        )
    return sorted(folders, key=lambda item: item["relative"].lower())


def normalize_workbook_name(name_value: str) -> str:
    name = clean_text(name_value) or DEFAULT_NEW_WORKBOOK_NAME
    for char in '<>:"/\\|?*':
        name = name.replace(char, "-")
    if Path(name).suffix.lower() not in EXCEL_SUFFIXES:
        name = f"{name}.xlsx"
    return name


def set_active_workbook(path_value: str) -> Path:
    raw = clean_text(path_value)
    if not raw:
        raise ValueError("Selecione uma planilha ou informe um caminho válido.")
    path = Path(raw)
    if not path.is_absolute():
        path = PROJECT_DIR / path
    path = path.resolve()
    if path.suffix.lower() not in EXCEL_SUFFIXES:
        raise ValueError("Use uma planilha .xlsx ou .xlsm.")
    if not path.parent.exists():
        raise FileNotFoundError(f"Pasta não encontrada: {path.parent}")
    config = _read_config()
    config["active_workbook"] = str(path)
    _write_config(config)
    return path


def set_active_workbook_from_folder(folder_value: str, workbook_name: str) -> Path:
    raw_folder = clean_text(folder_value)
    if not raw_folder:
        raise ValueError("Selecione uma pasta ou informe um caminho de pasta válido.")
    folder = Path(raw_folder)
    if not folder.is_absolute():
        folder = PROJECT_DIR / folder
    folder = folder.resolve()
    if not folder.exists() or not folder.is_dir():
        raise FileNotFoundError(f"Pasta não encontrada: {folder}")
    path = folder / normalize_workbook_name(workbook_name)
    ensure_workbook_exists(path)
    config = _read_config()
    config["active_workbook"] = str(path)
    _write_config(config)
    return path


def _safe_sheet_title(label: str) -> str:
    title = INVALID_SHEET_CHARS_RE.sub("-", clean_text(label) or "Categoria").strip("' ")
    return (title or "Categoria")[:31]


def _sheet_title_for_category(category: dict[str, Any]) -> str:
    return _safe_sheet_title(category.get("sheet_name") or category.get("label") or "Categoria")


def _sheet_for_category(workbook, category: dict[str, Any]):
    title = _sheet_title_for_category(category)
    if title in workbook.sheetnames:
        return workbook[title]
    ws = workbook.create_sheet(title=title)
    _write_headers(ws, get_banco_fields(category["key"]))
    return ws


def _headers_from_fields(fields: list[dict[str, Any]]) -> list[str]:
    headers = ["DESCRIÇÃO PRIMÁRIA", "DESCRIÇÃO SECUNDÁRIA"]
    for field in fields:
        header = header_for_field(field["label"], field["scope"])
        headers.append(header)
    return headers


def _style_header_cell(cell, fill_color: str) -> None:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=fill_color)
    cell.alignment = Alignment(horizontal="center", vertical="center")


def create_blank_workbook(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    default_sheet = workbook.active
    default_sheet.title = DEFAULT_CATEGORY_LABEL
    _write_headers(default_sheet, get_banco_fields(DEFAULT_CATEGORY_KEY))
    registration_sheet = workbook.create_sheet(REGISTRATION_SHEET_NAME)
    registration_sheet.sheet_state = "hidden"
    registration_sheet.append(["category_key", "category_label", "sheet", "row", "saved_at"])
    workbook.save(path)
    workbook.close()


def ensure_workbook_exists(path: Path | None = None) -> Path:
    workbook = (path or active_workbook_path()).resolve()
    if workbook.exists():
        return workbook
    if DEFAULT_TEMPLATE_PATH.exists():
        workbook.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(DEFAULT_TEMPLATE_PATH, workbook)
        return workbook
    create_blank_workbook(workbook)
    return workbook


def _write_headers(ws, fields: list[dict[str, Any]]) -> None:
    headers = _headers_from_fields(fields)
    for index, header in enumerate(headers, start=1):
        ws.cell(1, index).value = header
        _style_header_cell(ws.cell(1, index), "0F172A" if index <= 2 else "1D4ED8")
    for index in range(len(headers) + 1, ws.max_column + 1):
        ws.cell(1, index).value = None
    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 42
    for field_index, field in enumerate(fields):
        value_column = 3 + field_index
        ws.column_dimensions[get_column_letter(value_column)].width = max(20, len(field["label"]) + 4)


def _candidate_header_rows(ws) -> list[int]:
    rows: list[int] = []
    tables = getattr(ws, "tables", None) or {}
    for table in tables.values():
        ref = getattr(table, "ref", table)
        if isinstance(ref, str):
            _, min_row, _, _ = range_boundaries(ref)
            rows.append(min_row)
    rows.extend([1, 2, 3])

    unique_rows: list[int] = []
    for row in rows:
        if row >= 1 and row not in unique_rows and row <= max(ws.max_row, 3):
            unique_rows.append(row)
    return unique_rows


def _detect_header_row(ws) -> int:
    best_row = 1
    best_score = -1
    for row in _candidate_header_rows(ws):
        score = 0
        for column in range(1, ws.max_column + 1):
            if clean_text(ws.cell(row, column).value):
                score += 1
        if score > best_score:
            best_score = score
            best_row = row
    return best_row


def _table_for_sheet(ws):
    tables = getattr(ws, "tables", None) or {}
    for table in tables.values():
        return table
    return None


def _table_bounds(ws) -> tuple[int, int, int, int] | None:
    table = _table_for_sheet(ws)
    if table is None:
        return None
    ref = getattr(table, "ref", None)
    if not ref:
        return None
    return range_boundaries(ref)


def _row_has_values(ws, row: int, start_column: int, end_column: int) -> bool:
    for column in range(start_column, end_column + 1):
        if clean_text(ws.cell(row, column).value):
            return True
    return False


def _next_available_row(ws) -> int:
    primary_column, _, _ = _resolve_description_columns(ws, create_missing=False)
    if primary_column:
        bounds = _table_bounds(ws)
        if bounds is not None:
            _, min_row, _, max_row = bounds
            start_row = max(FIRST_DATA_ROW, min_row + 1)
            for row in range(start_row, max_row + 1):
                if not clean_text(ws.cell(row, primary_column).value):
                    return row
            return max_row + 1

        header_row = _detect_header_row(ws)
        start_row = max(FIRST_DATA_ROW, header_row + 1)
        row = start_row
        while clean_text(ws.cell(row, primary_column).value):
            row += 1
        return row

    header_row = _detect_header_row(ws)
    return max(FIRST_DATA_ROW, header_row + 1)


def _expand_table_to_row(ws, row: int) -> None:
    table = _table_for_sheet(ws)
    if table is None:
        return
    ref = getattr(table, "ref", None)
    if not ref:
        return
    min_col, min_row, max_col, max_row = range_boundaries(ref)
    if row <= max_row:
        return
    table.ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{row}"


def _code_header_for_field(field: dict[str, Any]) -> str:
    return f"{header_for_field(field['label'], field['scope'])} CÓDIGO"


def _worksheet_uses_code_columns(ws) -> bool:
    header_row = _detect_header_row(ws)
    for column in range(3, ws.max_column + 1):
        if "CODIGO" in normalize_label(ws.cell(header_row, column).value):
            return True
    return False


def _header_column_map(ws) -> dict[str, int]:
    header_row = _detect_header_row(ws)
    mapping: dict[str, int] = {}
    for column in range(1, ws.max_column + 1):
        normalized = normalize_label(ws.cell(header_row, column).value)
        if normalized and normalized not in mapping:
            mapping[normalized] = column
    return mapping


def _last_header_column(ws, header_row: int | None = None) -> int:
    header_row = header_row or _detect_header_row(ws)
    for column in range(ws.max_column, 0, -1):
        if clean_text(ws.cell(header_row, column).value):
            return column
    return 0


def _append_header_cell(ws, header: str, header_row: int | None = None) -> int:
    header_row = header_row or _detect_header_row(ws)
    column = _last_header_column(ws, header_row) + 1
    ws.cell(header_row, column).value = header
    _style_header_cell(ws.cell(header_row, column), "1D4ED8")
    ws.column_dimensions[get_column_letter(column)].width = max(20, len(clean_text(header)) + 4)
    return column


def _resolve_header_column(ws, header: str, create_missing: bool = False) -> int | None:
    header_row = _detect_header_row(ws)
    header_map = _header_column_map(ws)
    column = header_map.get(normalize_label(header))
    if column is not None or not create_missing:
        return column
    return _append_header_cell(ws, header, header_row)


def _resolve_description_columns(ws, create_missing: bool = False) -> tuple[int | None, int | None, int | None]:
    primary_column = _resolve_header_column(ws, DESCRIPTION_PRIMARY_HEADER, create_missing)
    secondary_column = _resolve_header_column(ws, DESCRIPTION_SECONDARY_HEADER, create_missing)
    suffix_column = _resolve_header_column(ws, DESCRIPTION_SUFFIX_HEADER, create_missing)
    return primary_column, secondary_column, suffix_column


def _resolve_field_columns(ws, field: dict[str, Any], create_missing: bool = False) -> tuple[int | None, int | None]:
    header_map = _header_column_map(ws)
    value_header = normalize_label(header_for_field(field["label"], field["scope"]))
    code_header = normalize_label(_code_header_for_field(field))

    value_column = header_map.get(value_header)
    code_column = header_map.get(code_header)

    uses_code_columns = _worksheet_uses_code_columns(ws)
    if value_column is None and create_missing:
        value_column = _append_header_cell(ws, header_for_field(field["label"], field["scope"]))
        if uses_code_columns:
            code_column = _append_header_cell(ws, _code_header_for_field(field))

    return value_column, code_column


def _resolve_field_column_map(
    ws,
    fields: list[dict[str, Any]],
    create_missing: bool = False,
) -> dict[str, tuple[int | None, int | None]]:
    header_map = _header_column_map(ws)
    uses_code_columns = any("CODIGO" in header for header in header_map)
    column_map: dict[str, tuple[int | None, int | None]] = {}

    for field in fields:
        value_header_text = header_for_field(field["label"], field["scope"])
        code_header_text = _code_header_for_field(field)
        value_header = normalize_label(value_header_text)
        code_header = normalize_label(code_header_text)

        value_column = header_map.get(value_header)
        code_column = header_map.get(code_header)

        if value_column is None and create_missing:
            value_column = _append_header_cell(ws, value_header_text)
            header_map[value_header] = value_column
            if uses_code_columns:
                code_column = _append_header_cell(ws, code_header_text)
                header_map[code_header] = code_column

        column_map[field["key"]] = (value_column, code_column)

    return column_map


def _save_workbook_atomic(wb, target_path: Path) -> None:
    target_path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix="_cadastro_app_save_", suffix=target_path.suffix or ".xlsx", dir=str(target_path.parent))
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        wb.save(temp_path)
        os.replace(temp_path, target_path)
    finally:
        temp_path.unlink(missing_ok=True)


def _save_workbook_preserving_package(wb, target_path: Path) -> None:
    _save_workbook_atomic(wb, target_path)


def _preserve_workbook_parts(source_path: Path, target_path: Path, prefixes: tuple[str, ...]) -> None:
    if not source_path.exists() or not target_path.exists():
        return

    source_members: dict[str, zipfile.ZipInfo] = {}
    source_payloads: dict[str, bytes] = {}
    with zipfile.ZipFile(source_path, "r") as source_zip:
        for member in source_zip.infolist():
            if any(member.filename.startswith(prefix) for prefix in prefixes):
                source_members[member.filename] = member
                source_payloads[member.filename] = source_zip.read(member.filename)

    if not source_payloads:
        return

    fd, temp_name = tempfile.mkstemp(prefix="_cadastro_app_zip_", suffix=target_path.suffix or ".xlsx", dir=str(target_path.parent))
    os.close(fd)
    temp_path = Path(temp_name)

    try:
        with zipfile.ZipFile(target_path, "r") as target_zip, zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as output_zip:
            target_names = set(target_zip.namelist())
            for member in target_zip.infolist():
                payload = source_payloads.get(member.filename, target_zip.read(member.filename))
                info = source_members.get(member.filename, member)
                output_zip.writestr(info, payload)
            for filename, info in source_members.items():
                if filename not in target_names:
                    output_zip.writestr(info, source_payloads[filename])
        os.replace(temp_path, target_path)
    finally:
        temp_path.unlink(missing_ok=True)


def _field_response(field: dict[str, Any], index: int) -> dict[str, Any]:
    column = 3 + index
    options = list(field.get("options") or [])
    return {
        "key": field["key"],
        "label": field["label"],
        "header": header_for_field(field["label"], field["scope"]),
        "scope": field["scope"],
        "selection_mode": field.get("selection_mode", SELECTION_MODE_UNITARIA),
        "column": column,
        "letter": get_column_letter(column),
        "options": options,
        "option_rows": [{"row": option_index, "value": value} for option_index, value in enumerate(options, start=1)],
    }


def get_banco_fields(category_key_value: str) -> list[dict[str, Any]]:
    category = _find_category(load_catalog(), category_key_value)
    return [_field_response(field, index) for index, field in enumerate(category.get("fields") or [])]


def _sheet_for_view(workbook, category: dict[str, Any]):
    title = _sheet_title_for_category(category)
    if title in workbook.sheetnames:
        return workbook[title]
    return None


def list_saved_registrations(category_key_value: str) -> dict[str, Any]:
    catalog = load_catalog()
    category = _find_category(catalog, category_key_value)
    selected = {"key": category["key"], "label": category["label"]}
    fields = get_banco_fields(category["key"])
    workbook = active_workbook_path()
    sheet_name = _sheet_title_for_category(category)
    result = {
        "category": selected,
        "fields": fields,
        "records": [],
        "workbook_path": str(workbook),
        "workbook_exists": workbook.exists(),
        "sheet_name": sheet_name,
        "sheet_exists": False,
        "workbook_locked": False,
        "error_message": "",
    }
    if not workbook.exists():
        return result

    workbook_mtime_ns = workbook.stat().st_mtime_ns
    cache_key = (str(workbook.resolve()), category["key"], workbook_mtime_ns)
    cached = _REGISTRATION_CACHE.get(cache_key)
    if cached is not None:
        return deepcopy(cached)

    try:
        wb = load_workbook(workbook, data_only=True)
    except PermissionError:
        result["workbook_locked"] = True
        result["error_message"] = (
            "A planilha está aberta em outro programa e não pode ser lida agora. "
            "Feche o Excel para visualizar os cadastros."
        )
        return result

    try:
        ws = _sheet_for_view(wb, category)
        if ws is None:
            return result

        primary_column, secondary_column, suffix_column = _resolve_description_columns(ws, create_missing=False)
        field_columns = _resolve_field_column_map(ws, fields, create_missing=False)
        result["sheet_name"] = ws.title
        result["sheet_exists"] = True
        bounds = _table_bounds(ws)
        last_row = bounds[3] if bounds else ws.max_row
        for row_index, row_values in enumerate(
            ws.iter_rows(min_row=FIRST_DATA_ROW, max_row=last_row, values_only=True),
            start=FIRST_DATA_ROW,
        ):
            descricao_primaria = clean_text(row_values[primary_column - 1]) if primary_column and primary_column <= len(row_values) else ""
            descricao_secundaria = compose_secondary_description(
                descricao_primaria,
                row_values[secondary_column - 1] if secondary_column and secondary_column <= len(row_values) else "",
            )
            sufixo = clean_text(row_values[suffix_column - 1]) if suffix_column and suffix_column <= len(row_values) else ""
            values: list[dict[str, str]] = []
            has_field_value = False

            for field_index, field in enumerate(fields):
                value_column, _ = field_columns.get(field["key"], (None, None))
                value = clean_text(row_values[value_column - 1]) if value_column and value_column <= len(row_values) else ""
                if value:
                    has_field_value = True
                values.append({"key": field["key"], "label": field["label"], "value": value})

            if not descricao_primaria and not descricao_secundaria and not has_field_value:
                continue

            result["records"].append(
                {
                    "row": row_index,
                    "descricao_primaria": descricao_primaria,
                    "descricao_secundaria": descricao_secundaria,
                    "sufixo": sufixo,
                    "values": values,
                }
            )
        _REGISTRATION_CACHE[cache_key] = deepcopy(result)
        return result
    finally:
        wb.close()

def sync_workbook_headers(workbook: Path, category_key_value: str) -> None:
    catalog = load_catalog()
    category = _find_category(catalog, category_key_value)
    fields = get_banco_fields(category["key"])
    wb = _load(workbook)
    try:
        ws = _sheet_for_category(wb, category)
        primary_column, secondary_column, suffix_column = _resolve_description_columns(ws, create_missing=True)
        ws.freeze_panes = "A2"
        if primary_column:
            ws.column_dimensions[get_column_letter(primary_column)].width = 42
        if secondary_column:
            ws.column_dimensions[get_column_letter(secondary_column)].width = 42
        if suffix_column:
            ws.column_dimensions[get_column_letter(suffix_column)].width = 18
        _resolve_field_column_map(ws, fields, create_missing=True)
        _registration_sheet(wb)
        _save_workbook_atomic(wb, workbook)
    except PermissionError as exc:
        raise PermissionError(
            "Não consegui salvar a planilha. Feche o Excel se ela estiver aberta e tente novamente."
        ) from exc
    finally:
        wb.close()


def sync_workbook_structure(category_key_value: str) -> None:
    sync_workbook_headers(active_workbook_path(), category_key_value)


def _registration_sheet(workbook):
    if REGISTRATION_SHEET_NAME not in workbook.sheetnames:
        ws = workbook.create_sheet(REGISTRATION_SHEET_NAME)
        ws.sheet_state = "hidden"
        ws.append(["category_key", "category_label", "sheet", "row", "saved_at"])
        return ws
    ws = workbook[REGISTRATION_SHEET_NAME]
    if ws.max_row == 1 and clean_text(ws.cell(1, 1).value) != "category_key":
        ws.delete_rows(1, ws.max_row)
        ws.append(["category_key", "category_label", "sheet", "row", "saved_at"])
    ws.sheet_state = "hidden"
    return ws


def _copy_row_style(ws, source_row: int, target_row: int) -> None:
    for column in range(1, ws.max_column + 1):
        source = ws.cell(source_row, column)
        target = ws.cell(target_row, column)
        if source.has_style:
            target._style = copy(source._style)
        if source.number_format:
            target.number_format = source.number_format
        if source.font:
            target.font = copy(source.font)
        if source.fill:
            target.fill = copy(source.fill)
        if source.border:
            target.border = copy(source.border)
        if source.alignment:
            target.alignment = copy(source.alignment)
        if source.protection:
            target.protection = copy(source.protection)
    if source_row in ws.row_dimensions:
        ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height


def _copy_row_formulas(ws, source_row: int, target_row: int) -> None:
    for column in range(1, ws.max_column + 1):
        source = ws.cell(source_row, column)
        target = ws.cell(target_row, column)
        value = source.value
        if isinstance(value, ArrayFormula):
            target.value = ArrayFormula(ref=target.coordinate, text=value.text)
        elif source.data_type == "f" or (isinstance(value, str) and value.startswith("=")):
            target.value = value


def _first_code(value: str, fallback: int) -> int:
    match = VALUE_RE.match(clean_text(value))
    if match:
        return int(match.group(1))
    return fallback


def _serialize_field_values(field: dict[str, Any], data: Any) -> list[str]:
    raw_values: list[str] = []
    key = field["key"]
    if hasattr(data, "getlist"):
        raw_values.extend(data.getlist(key))
    else:
        value = data.get(key)
        if isinstance(value, list):
            raw_values.extend(value)
        elif value is not None:
            raw_values.append(value)

    if field.get("selection_mode") != SELECTION_MODE_MULTIPLA:
        raw_values = raw_values[:1]

    cleaned_values = [clean_text(value) for value in raw_values if clean_text(value)]
    if not cleaned_values:
        return []

    options = field.get("options") or []
    submitted_identities = {option_identity(value): value for value in cleaned_values}
    ordered: list[str] = []
    used: set[str] = set()

    for option in options:
        identity = option_identity(option)
        if identity in submitted_identities and identity not in used:
            ordered.append(option)
            used.add(identity)

    for value in cleaned_values:
        identity = option_identity(value)
        if identity not in used:
            ordered.append(value)
            used.add(identity)

    return ordered


def _format_distancia_pe_value(values: list[str]) -> str:
    joined = " | ".join(values).strip()
    if not joined:
        return ""

    normalized_joined = normalize_label(joined)
    normalized_prefix = normalize_label(DISTANCIA_PE_PREFIX)
    if normalized_joined.startswith(normalized_prefix):
        return joined
    return f"{DISTANCIA_PE_PREFIX} {joined}".strip()


def _format_field_saved_value(field: dict[str, Any], values: list[str]) -> str:
    if field.get("key") == DISTANCIA_PE_KEY:
        return _format_distancia_pe_value(values)
    return " | ".join(values)


def _format_field_description(field: dict[str, Any], values: list[str]) -> str:
    label = " / ".join(option_label(value) for value in values)
    if field.get("key") == DISTANCIA_PE_KEY and label:
        prefix = DISTANCIA_PE_PREFIX
        if normalize_label(label).startswith(normalize_label(prefix)):
            return label
        return f"{prefix} {label}".strip()
    return label


def build_descriptions(fields: list[dict[str, Any]], data: Any) -> dict[str, str]:
    primary_parts: list[str] = []
    secondary_parts: list[str] = []
    secondary_codes: list[str] = []

    for field in fields:
        values = _serialize_field_values(field, data)
        if not values:
            continue
        label = _format_field_description(field, values)
        if field["scope"] == "primaria":
            primary_parts.append(label)
        elif field["scope"] == "secundaria":
            secondary_parts.append(label)
            secondary_codes.extend(code for code in (option_code(value) for value in values) if code)

    primary_description_base = " ".join(primary_parts).strip()
    suffix = ""
    if secondary_codes:
        suffix = ".".join(secondary_codes)
    secondary_description = " ".join(secondary_parts).strip()

    return {
        "primaria": primary_description_base,
        "secundaria": compose_secondary_description(primary_description_base, secondary_description),
        "sufixo": suffix,
    }


def _submitted_values(fields: list[dict[str, Any]], data: Any) -> dict[str, str]:
    return {field["key"]: " | ".join(_serialize_field_values(field, data)) for field in fields}


def _field_values_by_key(fields: list[dict[str, Any]], data: Any) -> dict[str, list[str]]:
    return {field["key"]: _serialize_field_values(field, data) for field in fields}


def _has_any_option_code(values: list[str], allowed_codes: set[str]) -> bool:
    return any(option_code(value) in allowed_codes for value in values)


def _validate_banco_dependencies(fields: list[dict[str, Any]], data: Any) -> None:
    values_by_key = _field_values_by_key(fields, data)
    pre_fixo_values = values_by_key.get("pre_fixo", [])
    veiculo_values = values_by_key.get("veiculo", [])

    if _has_any_option_code(pre_fixo_values, {"4", "5", "6"}) and not veiculo_values:
        raise ValueError(
            "O campo VEÍCULO é obrigatório quando PRE FIXO for "
            "BCO CARONA ORIGINAL, BCO MOTORISTA ORIGINAL ou BCO ORIGINAL."
        )


def _has_any_value(values: dict[str, str]) -> bool:
    return any(values.values())


def _row_values(ws, field_columns: dict[str, tuple[int | None, int | None]], row: int) -> dict[str, str]:
    values: dict[str, str] = {}
    for field_key, (value_column, _) in field_columns.items():
        values[field_key] = clean_text(ws.cell(row, value_column).value) if value_column else ""
    return values


def _find_duplicate_registration(
    ws,
    fields: list[dict[str, Any]],
    field_columns: dict[str, tuple[int | None, int | None]],
    data: Any,
) -> int | None:
    submitted = _submitted_values(fields, data)
    if not _has_any_value(submitted):
        raise ValueError("Selecione pelo menos uma opção antes de salvar o cadastro.")

    for row in range(FIRST_DATA_ROW, ws.max_row + 1):
        existing = _row_values(ws, field_columns, row)
        if not _has_any_value(existing):
            continue
        if existing == submitted:
            return row
    return None


def _backup_workbook(workbook: Path, suffix: str) -> Path:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = BACKUP_DIR / f"{workbook.stem}_{suffix}_{stamp}{workbook.suffix}"
    shutil.copy2(workbook, backup_path)
    return backup_path


def save_banco_registration(form_data: Any) -> dict[str, str]:
    category_key_value = clean_text(form_data.get("categoria"))
    category = selected_category(category_key_value)
    workbook = ensure_workbook_exists()
    workbook_source = _copy_to_temp(workbook)
    wb = _load(workbook_source)
    backup_path = None

    try:
        catalog = load_catalog()
        raw_category = _find_category(catalog, category["key"])
        ws = _sheet_for_category(wb, raw_category)
        fields = get_banco_fields(category["key"])
        _validate_banco_dependencies(fields, form_data)
        primary_column, secondary_column, suffix_column = _resolve_description_columns(ws, create_missing=True)
        field_columns = _resolve_field_column_map(ws, fields, create_missing=True)
        duplicate_row = _find_duplicate_registration(ws, fields, field_columns, form_data)
        if duplicate_row:
            raise ValueError(f"Cadastro já existe na linha {duplicate_row}.")

        backup_path = _backup_workbook(workbook, "cadastro")
        row = _next_available_row(ws)
        if row > FIRST_DATA_ROW:
            _copy_row_style(ws, row - 1, row)
            _copy_row_formulas(ws, row - 1, row)
        _expand_table_to_row(ws, row)

        descriptions = build_descriptions(fields, form_data)
        if primary_column:
            ws.cell(row, primary_column).value = descriptions["primaria"]
        if secondary_column:
            ws.cell(row, secondary_column).value = descriptions["secundaria"]
        if suffix_column:
            ws.cell(row, suffix_column).value = descriptions["sufixo"] or None

        for field in fields:
            values = _serialize_field_values(field, form_data)
            value_column, code_column = field_columns.get(field["key"], (None, None))
            ws.cell(row, value_column).value = _format_field_saved_value(field, values) if values else None
            if code_column:
                codes = [option_code(value) for value in values if option_code(value)]
                ws.cell(row, code_column).value = " | ".join(codes) if codes else None

        control_ws = _registration_sheet(wb)
        control_ws.append(
            [
                raw_category["key"],
                raw_category["label"],
                ws.title,
                row,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ]
        )

        _save_workbook_preserving_package(wb, workbook)
        _REGISTRATION_CACHE.clear()
        set_active_category(category["key"])
        return {
            "row": row,
            "category": raw_category["label"],
            "category_key": raw_category["key"],
            "sheet": ws.title,
            "descricao_primaria": descriptions["primaria"],
            "descricao_secundaria": descriptions["secundaria"],
            "path": str(workbook),
            "backup": str(backup_path) if backup_path else "",
        }
    except PermissionError as exc:
        raise PermissionError(
            "Não consegui salvar a planilha. Feche o Excel se ela estiver aberta e tente novamente."
        ) from exc
    finally:
        wb.close()
        cleanup_path = getattr(wb, "_cadastro_cleanup_path", None)
        if cleanup_path is not None:
            try:
                Path(cleanup_path).unlink(missing_ok=True)
            except OSError:
                pass
        try:
            workbook_source.unlink(missing_ok=True)
        except OSError:
            pass


def _find_field(catalog: dict[str, Any], category_key_value: str, field_key_value: str) -> dict[str, Any]:
    category = _find_category(catalog, category_key_value)
    for field in category.get("fields") or []:
        if field["key"] == field_key_value:
            return field
    raise ValueError("Campo não encontrado.")


def _next_option_code(options: list[str]) -> int:
    max_code = 0
    for option in options:
        match = VALUE_RE.match(clean_text(option))
        if match:
            max_code = max(max_code, int(match.group(1)))
    return max_code + 1


def _format_option_value(value: str, options: list[str], existing_value: str = "") -> str:
    text = clean_text(value)
    if not text:
        return ""
    if VALUE_RE.match(text):
        return normalize_option_text(text)
    code = _first_code(existing_value, _next_option_code(options))
    return normalize_option_text(f"{code}- {text}")


def add_field_option(category_key_value: str, field_key_value: str, option_value: str) -> dict[str, str]:
    catalog = load_catalog()
    field = _find_field(catalog, category_key_value, field_key_value)
    raw_option = clean_text(option_value)
    if not raw_option:
        raise ValueError("Informe a nova opção.")

    options = field.setdefault("options", [])
    if option_identity(raw_option) in {option_identity(value) for value in options}:
        raise ValueError("Essa opção já existe para o campo selecionado.")

    option = _format_option_value(raw_option, options)
    options.append(option)
    save_catalog(catalog)
    return {
        "field": field["label"],
        "option": option,
        "row": len(options),
        "path": str(DATA_PATH),
        "backup": "",
    }


def update_field_option(category_key_value: str, field_key_value: str, row_value: int, option_value: str) -> dict[str, str]:
    catalog = load_catalog()
    field = _find_field(catalog, category_key_value, field_key_value)
    options = field.setdefault("options", [])
    index = int(row_value) - 1
    if index < 0 or index >= len(options):
        raise ValueError("Opção não encontrada.")

    current = options[index]
    option = _format_option_value(option_value, options, current)
    other_options = [value for pos, value in enumerate(options) if pos != index]
    if option_identity(option) in {option_identity(value) for value in other_options}:
        raise ValueError("Essa opção já existe para o campo selecionado.")

    options[index] = option
    save_catalog(catalog)
    return {
        "field": field["label"],
        "option": option,
        "row": row_value,
        "path": str(DATA_PATH),
        "backup": "",
    }


def update_field_options(
    category_key_value: str,
    field_key_value: str,
    row_values: list[int],
    option_values: list[str],
) -> dict[str, Any]:
    catalog = load_catalog()
    field = _find_field(catalog, category_key_value, field_key_value)
    options = field.setdefault("options", [])
    if len(row_values) != len(option_values):
        raise ValueError("Quantidade de alterações inválida.")

    updates: list[tuple[int, str, int]] = []
    for row_value, raw_option in zip(row_values, option_values):
        index = int(row_value) - 1
        if index < 0 or index >= len(options):
            raise ValueError("Opção não encontrada.")
        raw_text = clean_text(raw_option)
        if not raw_text:
            continue
        current = options[index]
        option = _format_option_value(raw_text, options, current)
        updates.append((index, option, row_value))

    if not updates:
        return {
            "field": field["label"],
            "updated": [],
            "count": 0,
            "path": str(DATA_PATH),
            "backup": "",
        }

    new_options = list(options)
    for index, option, _ in updates:
        new_options[index] = option

    seen_identities: set[str] = set()
    for option in new_options:
        identity = option_identity(option)
        if identity in seen_identities:
            raise ValueError("Essa opção já existe para o campo selecionado.")
        seen_identities.add(identity)

    options[:] = new_options
    save_catalog(catalog)
    return {
        "field": field["label"],
        "updated": [option for _, option, _ in updates],
        "count": len(updates),
        "path": str(DATA_PATH),
        "backup": "",
    }


def delete_field_option(category_key_value: str, field_key_value: str, row_value: int) -> dict[str, str]:
    catalog = load_catalog()
    field = _find_field(catalog, category_key_value, field_key_value)
    options = field.setdefault("options", [])
    index = int(row_value) - 1
    if index < 0 or index >= len(options):
        raise ValueError("Opção não encontrada.")
    option = options.pop(index)
    save_catalog(catalog)
    return {
        "field": field["label"],
        "option": option,
        "row": row_value,
        "path": str(DATA_PATH),
        "backup": "",
    }


def add_field(category_key_value: str, label: str, scope: str, selection_mode: str) -> dict[str, str]:
    catalog = load_catalog()
    category = _find_category(catalog, category_key_value)
    label_clean = clean_text(label).upper()
    if not label_clean:
        raise ValueError("Informe o nome do campo.")
    if normalize_label(label_clean) in {normalize_label(field["label"]) for field in category["fields"]}:
        raise ValueError("Esse campo já existe nesta categoria.")

    used = {field["key"] for field in category["fields"]}
    field = {
        "key": _unique_key(label_clean, used),
        "label": label_clean,
        "scope": field_scope(scope),
        "selection_mode": field_selection_mode(selection_mode),
        "options": [],
    }
    category["fields"].append(field)
    save_catalog(catalog)
    sync_workbook_structure(category["key"])
    return {
        "field": field["label"],
        "scope": field["scope"],
        "selection_mode": field["selection_mode"],
        "path": str(DATA_PATH),
        "backup": "",
    }


def update_field(
    category_key_value: str,
    field_key_value: str,
    label: str,
    scope: str,
    selection_mode: str,
) -> dict[str, str]:
    catalog = load_catalog()
    category = _find_category(catalog, category_key_value)
    field = _find_field(catalog, category["key"], field_key_value)
    label_clean = clean_text(label).upper()
    if not label_clean:
        raise ValueError("Informe o nome do campo.")

    for existing in category["fields"]:
        if existing["key"] != field_key_value and normalize_label(existing["label"]) == normalize_label(label_clean):
            raise ValueError("Já existe outro campo com esse nome nesta categoria.")

    field["label"] = label_clean
    field["scope"] = field_scope(scope)
    field["selection_mode"] = field_selection_mode(selection_mode)
    save_catalog(catalog)
    sync_workbook_structure(category["key"])
    return {
        "field": field["label"],
        "scope": field["scope"],
        "selection_mode": field["selection_mode"],
        "path": str(DATA_PATH),
        "backup": "",
    }


def delete_field(category_key_value: str, field_key_value: str) -> dict[str, str]:
    catalog = load_catalog()
    category = _find_category(catalog, category_key_value)
    field = _find_field(catalog, category["key"], field_key_value)
    category["fields"] = [item for item in category["fields"] if item["key"] != field_key_value]
    save_catalog(catalog)
    return {"field": field["label"], "path": str(DATA_PATH), "backup": ""}
