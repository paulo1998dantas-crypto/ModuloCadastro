"""Build a reviewed recovery payload for the 2026-08-25 Banks incident.

The source workbook preserved all 347 active SKUs, but its visible headers were
shifted relative to the technical keys.  This tool maps columns explicitly,
canonicalizes every selectable value against the current Cadastro catalogue and
emits data only; it never connects to or changes the database.
"""

from __future__ import annotations

import hashlib
import json
import re
import sys
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "outputs" / "bank_fields_source_20260825.json"
OUTPUT = ROOT / "outputs" / "bank_recovery_payload_20260825.json"
sys.path.insert(0, str(ROOT))

import excel_bancos
import supabase_store
import technical_fields_reconciliation as reconciliation


# Target technical key -> zero-based source column.  The source workbook was
# recovered after a stale header projection permuted these columns.
COLUMN_MAP = {
    "pre_fixo": 3,
    "fornecedor": 4,
    "linha": 5,
    "encosto": 6,
    "cj_layout": 7,
    "lotacao": 8,
    "especificidade": 9,
    "braco": 10,
    "lado_braco": 11,
    "cj_acessibilidade": 12,
    "altura_pe": 13,
    "tipo_cinto": 14,
    "tipo_revestimento": 15,
    "tipo_da_espuma": 16,
    "largura_encosto": 17,
    "cj_observacao": 18,
    "profundidade_assento": 19,
    "tipo_apoio_braco": 20,
    "posicao_braco": 21,
    "acabamento_lateral": 22,
    "usb": 23,
    "resfriador": 24,
    "aquecedor": 25,
    "apoio_pe": 26,
    "massageador": 27,
    "mesa_snack": 28,
    "isofix": 29,
    "posicao_isofix": 30,
    "apoio_panturrilha": 31,
    "pega_mao": 32,
    "modelo_pe": 33,
    "qtd_pe": 34,
    "distancia_pe": 35,
    "tipo_do_reclinador": 36,
    "grau_reclinacao": 37,
    "veiculo": 38,
    "tipo_costura": 39,
    "cor_da_linha": 40,
    "cor_do_revestimento": 41,
}

SPECIAL_DESCRIPTIONS = {
    "10200172": "BCO SUCATA",
    "10200173": "BCO PROTOTIPO",
}


def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _source_value(row: list[Any], field_key: str) -> str:
    raw = _text(row[COLUMN_MAP[field_key]])
    if field_key == "cj_observacao" and raw in {"[]", "[ ]"}:
        return ""
    return raw


def _group_for_sku(sku: str) -> str:
    return sku[:2] if len(sku) >= 2 and sku[:2] in {"10", "20", "30"} else ""


def _recovery_canonical_values(field: dict[str, Any], raw_value: str, sku: str) -> list[str]:
    """Resolve only existing options, preserving meaningful parenthesized text.

    The generic reconciler intentionally removes parenthetical annotations while
    normalizing labels.  Banks contains canonical values such as
    ``(2REC / 1 REB)`` and ``(2FIX / 1REB)``; removing their contents makes them
    indistinguishable.  Recovery therefore uses the stricter option normalizer
    and never creates a new catalogue option.
    """
    parts = [part.strip() for part in re.split(r"\s*\|\s*", raw_value) if part.strip()]
    if field.get("selection_mode") != excel_bancos.SELECTION_MODE_MULTIPLA and len(parts) > 1:
        raise ValueError(f"SKU {sku}: o campo {field.get('label')} aceita somente uma opção.")
    if field.get("free_text"):
        return parts

    options = list(field.get("options") or [])
    canonical: list[str] = []
    for part in parts:
        expected = excel_bancos.normalize_option_label(part)
        matches = [
            option
            for option in options
            if excel_bancos.normalize_option_label(excel_bancos.option_label(option)) == expected
        ]
        if not matches:
            compact = re.sub(r"[^A-Z0-9]+", "", expected)
            matches = [
                option
                for option in options
                if re.sub(
                    r"[^A-Z0-9]+",
                    "",
                    excel_bancos.normalize_option_label(excel_bancos.option_label(option)),
                )
                == compact
            ]
        if len(matches) != 1:
            suggestions = sorted(
                (
                    SequenceMatcher(
                        None,
                        expected,
                        excel_bancos.normalize_option_label(excel_bancos.option_label(option)),
                    ).ratio(),
                    option,
                )
                for option in options
            )[-3:]
            raise ValueError(
                f"SKU {sku}: valor {part!r} não possui correspondência única no campo "
                f"{field.get('label')}. Sugestões: {[option for _, option in reversed(suggestions)]}"
            )
        canonical.append(matches[0])
    return canonical


def main() -> None:
    source_bytes = SOURCE.read_bytes()
    workbook = json.loads(source_bytes.decode("utf-8"))
    source_rows = workbook["CAMPOS_TECNICOS"][1:]
    fields = excel_bancos.get_banco_fields("bancos")
    fields_by_key = {field["key"]: field for field in fields}
    # The production catalogue gained these options after the checked-out
    # baseline; use the exact production codes without modifying the local
    # catalogue or creating anything during recovery.
    recovered_distance_options = [
        "64- SEGUNDO VAO 320 MM",
        "65- PRIMEIRO VAO 90 MM",
        "66- SEGUNDO VAO 160 MM",
        "67- PRIMEIRO VAO 190 MM",
        "68- SEGUNDO VAO 730 MM",
        "69- PRIMEIRO VAO 370 MM",
        "70- TERCEIRO VAO 1120 MM",
        "71- PRIMEIRO VAO 180 MM",
        "72- SEGUNDO VAO 830 MM",
        "73- PRIMEIRO VAO 100 MM",
        "74- SEGUNDO VAO 780 MM",
        "75- SEGUNDO VAO 570 MM",
        "76- QUARTO VAO 1410 MM",
        "77- SEGUNDO VAO 580 MM",
        "78- TERCEIRO VAO 1045 MM",
        "79- QUARTO VAO 1440 MM",
        "80- TERCEIRO VAO 1160 MM",
    ]
    current_distance_options = fields_by_key["distancia_pe"].setdefault("options", [])
    current_distance_labels = {
        excel_bancos.normalize_option_label(excel_bancos.option_label(option))
        for option in current_distance_options
    }
    current_distance_options.extend(
        option
        for option in recovered_distance_options
        if excel_bancos.normalize_option_label(excel_bancos.option_label(option))
        not in current_distance_labels
    )
    missing = sorted(set(fields_by_key) - set(COLUMN_MAP))
    extra = sorted(set(COLUMN_MAP) - set(fields_by_key))
    if missing or extra:
        raise RuntimeError(f"Mapeamento incompatível com o catálogo: ausentes={missing}; extras={extra}")

    unresolved: dict[str, set[str]] = {}
    for row in source_rows:
        sku = _text(row[0])
        if sku in SPECIAL_DESCRIPTIONS:
            continue
        for field in fields:
            raw = _source_value(row, field["key"])
            if field["key"] == "distancia_pe":
                raw = re.sub(r"^ORIENTADO\s+A\s+ESQ\s*:\s*", "", raw, flags=re.IGNORECASE)
            if raw and not field.get("free_text"):
                raw = " | ".join(
                    excel_bancos.option_label(part.strip()) or part.strip()
                    for part in raw.split("|")
                    if part.strip()
                )
            try:
                _recovery_canonical_values(field, raw, sku)
            except ValueError:
                unresolved.setdefault(field["key"], set()).add(raw)
    if unresolved:
        raise RuntimeError(
            "Valores sem correspondência exata no catálogo: "
            + json.dumps({key: sorted(values) for key, values in unresolved.items()}, ensure_ascii=False)
        )

    records: list[dict[str, Any]] = []
    source_skus: set[str] = set()
    for row in source_rows:
        sku = _text(row[0])
        if not sku or sku in source_skus:
            raise RuntimeError(f"SKU vazio ou duplicado na fonte: {sku!r}")
        source_skus.add(sku)
        group = _group_for_sku(sku)
        canonical: dict[str, list[str]] = {excel_bancos.PN_GROUP_FORM_KEY: [group] if group else []}
        if sku in SPECIAL_DESCRIPTIONS:
            canonical.update({field["key"]: [] for field in fields})
            descriptions = {
                "primaria": SPECIAL_DESCRIPTIONS[sku],
                "secundaria": SPECIAL_DESCRIPTIONS[sku],
                "sufixo": "",
            }
        else:
            for field in fields:
                raw = _source_value(row, field["key"])
                if field["key"] == "distancia_pe":
                    raw = re.sub(r"^ORIENTADO\s+A\s+ESQ\s*:\s*", "", raw, flags=re.IGNORECASE)
                if raw and not field.get("free_text"):
                    raw = " | ".join(
                        excel_bancos.option_label(part.strip()) or part.strip()
                        for part in raw.split("|")
                        if part.strip()
                    )
                canonical[field["key"]] = _recovery_canonical_values(field, raw, sku)
            canonical = supabase_store._field_groups(fields, canonical)
            descriptions = excel_bancos.build_descriptions(fields, canonical, "bancos")
            if not descriptions["primaria"] or not descriptions["secundaria"]:
                raise RuntimeError(f"SKU {sku}: descrição vazia após a recuperação")

        field_values = supabase_store._field_values(fields, canonical)
        field_codes = supabase_store._field_codes(fields, canonical)
        records.append(
            {
                "sku": sku,
                "technical_values": canonical,
                "descricao_primaria": descriptions["primaria"],
                "descricao_secundaria": descriptions["secundaria"],
                "sufixo": descriptions.get("sufixo") or "",
                "field_values": field_values,
                "field_codes": field_codes,
            }
        )

    payload = {
        "metadata": {
            "category_key": "bancos",
            "category_label": "20 - BANCOS",
            "source_sha256": hashlib.sha256(source_bytes).hexdigest(),
            "record_count": len(records),
            "field_count": len(fields),
            "technical_keys": sorted(COLUMN_MAP),
            "special_fallback_skus": sorted(SPECIAL_DESCRIPTIONS),
        },
        "records": records,
    }
    OUTPUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(payload["metadata"], ensure_ascii=False, indent=2))
    print(f"output={OUTPUT}")


if __name__ == "__main__":
    main()
