import io
import unittest
from copy import deepcopy
from unittest.mock import patch

from openpyxl import Workbook

import excel_bancos
import revestimento_reconciliation as reconciliation


def _fields():
    fields = []
    for index, field_key in enumerate(reconciliation.WORKBOOK_COLUMNS.values(), start=1):
        fields.append(
            {
                "key": field_key,
                "label": field_key.upper(),
                "scope": "primaria",
                "selection_mode": "unitaria",
                "description_order": index,
                "required": False,
                "free_text": False,
                "options": ["1- N/A"],
            }
        )
    next(field for field in fields if field["key"] == "descritor_base")["options"].append("2- ARO JAN")
    return fields


def _workbook_bytes(sku="10180001", category="18 - REVESTIMENTO", group="10 - INSUMO"):
    workbook = Workbook()
    worksheet = workbook.active
    headers = list(reconciliation.REQUIRED_COLUMNS)
    worksheet.append(["RELATORIO"])
    worksheet.append([])
    worksheet.append(headers)
    values = {
        "COD": sku,
        "CATEGORIA": category,
        "GRUPO": group,
        "ESTAGIO": "N/A",
        "IDENTIFICACAO": "ARO JAN.",
        "NIVEL": "N/A",
        "LOCAL": "N/A",
        "LADO": "N/A",
        "VEICULO": "N/A",
        "FORNECEDOR": "N/A",
        "TIPO": "N/A",
        "MATERIAL": "N/A",
        "ACABAMENTO": "N/A",
        "COR": "N/A",
        "ESPECIFICIDADE": "N/A",
    }
    worksheet.append([values[header] for header in headers])
    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


class RevestimentoReconciliationTests(unittest.TestCase):
    def test_equivalencia_de_pontuacao_reutiliza_opcao_existente(self):
        missing = reconciliation.missing_field_options(_workbook_bytes(), _fields())
        self.assertNotIn("descritor_base", missing)

    def test_validacao_impede_planilha_com_sku_incompleto(self):
        fields = _fields()
        active = [
            {"id": "1", "sku": "10180001", "form_values": {}},
            {"id": "2", "sku": "10180002", "form_values": {}},
        ]
        with self.assertRaisesRegex(ValueError, "ativos ausentes"):
            reconciliation.prepare_reconciliation(
                _workbook_bytes(),
                fields,
                active,
                lambda row, groups: {"id": row["id"], "form_values": groups},
            )

    def test_adicao_em_lote_salva_o_catalogo_uma_unica_vez(self):
        catalog = {
            "version": 2,
            "active_category": "cat_18_revestimento",
            "pn_groups": [],
            "categories": [
                {
                    "key": "cat_18_revestimento",
                    "label": "18 - REVESTIMENTO",
                    "sheet_name": "18 - REVESTIMENTO",
                    "fields": [
                        {
                            "key": "fornecedor",
                            "label": "FORNECEDOR",
                            "scope": "primaria",
                            "selection_mode": "unitaria",
                            "description_order": 1,
                            "required": False,
                            "free_text": False,
                            "options": ["1- JI"],
                        },
                        {
                            "key": "tipo",
                            "label": "TIPO",
                            "scope": "primaria",
                            "selection_mode": "unitaria",
                            "description_order": 2,
                            "required": False,
                            "free_text": False,
                            "options": ["1- ESSENCIAL"],
                        },
                    ],
                    "conditional_rules": [],
                }
            ],
        }
        saved = []
        with (
            patch.object(excel_bancos, "load_catalog", return_value=deepcopy(catalog)),
            patch.object(excel_bancos, "save_catalog", side_effect=lambda value: saved.append(deepcopy(value))),
        ):
            result = excel_bancos.add_category_field_options(
                "cat_18_revestimento",
                {"fornecedor": ["PILAR"], "tipo": ["PLUS"]},
            )

        self.assertEqual(result, {"fornecedor": ["2- PILAR"], "tipo": ["2- PLUS"]})
        self.assertEqual(len(saved), 1)


if __name__ == "__main__":
    unittest.main()
