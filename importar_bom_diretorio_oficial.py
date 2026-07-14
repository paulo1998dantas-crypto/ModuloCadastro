from __future__ import annotations

import json
from pathlib import Path

import supabase_store


DEFAULT_BOM_DIR_NAME = "02 - B.O.M"
DEFAULT_PROJECT_ROOT = Path(r"C:\Users\PRODUCAO-2.0\J I MONTADORA DE VEICULOS ESPECIAIS LTDA")


def find_default_bom_dir() -> Path:
    for path in DEFAULT_PROJECT_ROOT.rglob(DEFAULT_BOM_DIR_NAME):
        if path.is_dir():
            return path
    raise SystemExit(f"Diretorio {DEFAULT_BOM_DIR_NAME!r} nao encontrado em {DEFAULT_PROJECT_ROOT}")


def main() -> None:
    directory = find_default_bom_dir()
    result = supabase_store.import_bom_directory(directory)
    print(json.dumps({"directory": str(directory), **result}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
