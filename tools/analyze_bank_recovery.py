from __future__ import annotations

import json
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "outputs" / "bank_fields_source_20260825.json"
sys.path.insert(0, str(ROOT))

import excel_bancos


def normalized(value: object) -> str:
    return excel_bancos.normalize_label(excel_bancos.option_label(str(value or "")))


def main() -> None:
    data = json.loads(SOURCE.read_text(encoding="utf-8"))
    sheet = data["CAMPOS_TECNICOS"]
    headers = sheet[0]
    rows = sheet[1:]
    fields = excel_bancos.get_banco_fields("bancos")
    option_sets = {
        field["key"]: {normalized(option) for option in field.get("options", []) if normalized(option)}
        for field in fields
    }
    print(f"rows={len(rows)} fields={len(fields)} source_columns={len(headers) - 3}")
    for column_index, header in enumerate(headers[3:], start=3):
        values = [str(row[column_index]).strip() for row in rows if column_index < len(row) and row[column_index] not in (None, "")]
        unique = Counter(values)
        if not values:
            print(f"{column_index + 1:02d} {header!r}: empty")
            continue
        ranked: list[tuple[float, int, str]] = []
        for field in fields:
            if field.get("free_text"):
                continue
            matches = sum(1 for value in values if normalized(value) in option_sets[field["key"]])
            ranked.append((matches / len(values), matches, field["key"]))
        ranked.sort(reverse=True)
        best = ", ".join(f"{key}:{ratio:.1%}({count})" for ratio, count, key in ranked[:4])
        sample = " | ".join(value for value, _ in unique.most_common(5))
        print(f"{column_index + 1:02d} {header!r}: n={len(values)} unique={len(unique)} best=[{best}] sample=[{sample}]")

    print("\nOUTLIERS")
    checks = {
        "PRÉ-FIXO": "pre_fixo",
        "COR DO REVESTIMENTO": "posicao_braco",
    }
    for source_header, target_key in checks.items():
        column_index = headers.index(source_header)
        field = next(item for item in fields if item["key"] == target_key)
        unmatched = []
        for row in rows:
            sku = str(row[0]).strip()
            raw = str(row[column_index]).strip() if column_index < len(row) and row[column_index] not in (None, "") else ""
            if not raw:
                continue
            parts = [part.strip() for part in raw.split("|") if part.strip()]
            misses = [part for part in parts if normalized(part) not in option_sets[target_key]]
            if misses:
                unmatched.append((sku, raw, misses))
        print(source_header, "->", target_key, unmatched)


if __name__ == "__main__":
    main()
