import io
import unittest

from openpyxl import Workbook

import excel_bancos
import technical_fields_reconciliation as reconciliation


CATEGORY = {"key": "cat_12_vidros", "label": "12 - VIDROS"}


def _fields():
    return [
        {
            "key": "fornecedor",
            "label": "FORNECEDOR",
            "scope": "primaria",
            "selection_mode": excel_bancos.SELECTION_MODE_UNITARIA,
            "description_order": 1,
            "required": False,
            "free_text": False,
            "options": ["1- SALT"],
        },
        {
            "key": "especificidade",
            "label": "ESPECIFICIDADE",
            "scope": "secundaria",
            "selection_mode": excel_bancos.SELECTION_MODE_MULTIPLA,
            "description_order": 2,
            "required": False,
            "free_text": False,
            "options": ["1- N/A", "2- LATERAL"],
        },
    ]


def _rows():
    return [
        {
            "id": "item-1",
            "sku": "12180001",
            "form_values": {
                excel_bancos.PN_GROUP_FORM_KEY: ["10 - INSUMO"],
                "fornecedor": ["1- SALT"],
                "especificidade": ["1- N/A"],
            },
        },
        {
            "id": "item-2",
            "sku": "12180002",
            "form_values": {
                excel_bancos.PN_GROUP_FORM_KEY: ["10 - INSUMO"],
                "fornecedor": ["1- SALT"],
                "especificidade": ["2- LATERAL"],
            },
        },
    ]


def _workbook_bytes(rows):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "CAMPOS_TECNICOS"
    worksheet.append(["COD", "CATEGORIA", "GRUPO", "FORNECEDOR", "ESPECIFICIDADE"])
    for row in rows:
        worksheet.append(row)
    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


class TechnicalFieldsReconciliationTests(unittest.TestCase):
    def test_template_contains_all_active_skus_and_field_headers(self):
        content = reconciliation.build_template(CATEGORY, _fields(), _rows())
        loaded = reconciliation.load_rows(content, CATEGORY, _fields())
        self.assertEqual([row["sku"] for row in loaded], ["12180001", "12180002"])
        self.assertEqual(loaded[0]["values"]["fornecedor"], "SALT")

    def test_import_requires_exact_active_sku_set(self):
        content = _workbook_bytes([["12180001", "12 - VIDROS", "10 - INSUMO", "SALT", "N/A"]])
        with self.assertRaisesRegex(ValueError, "ativos ausentes"):
            reconciliation.prepare_reconciliation(
                content,
                CATEGORY,
                _fields(),
                _rows(),
                lambda row, values: {"id": row["id"], "form_values": values},
            )

    def test_new_option_is_previewed_and_multiple_values_are_canonicalized(self):
        content = _workbook_bytes(
            [
                ["12180001", "12 - VIDROS", "10 - INSUMO", "VTRX", "N/A | LATERAL"],
                ["12180002", "12 - VIDROS", "10 - INSUMO", "SALT", "LATERAL"],
            ]
        )
        additions = reconciliation.missing_field_options(content, CATEGORY, _fields())
        self.assertEqual(additions, {"fornecedor": ["VTRX"]})
        preview = reconciliation.fields_with_pending_options(_fields(), additions)
        result = reconciliation.prepare_reconciliation(
            content,
            CATEGORY,
            preview,
            _rows(),
            lambda row, values: {"id": row["id"], "form_values": values},
        )
        self.assertEqual(result["total"], 2)
        self.assertEqual(result["payloads"][0]["form_values"]["fornecedor"], ["2- VTRX"])
        self.assertEqual(
            result["payloads"][0]["form_values"]["especificidade"],
            ["1- N/A", "2- LATERAL"],
        )

    def test_group_is_informative_and_preserved_from_registration(self):
        content = _workbook_bytes(
            [
                ["12180001", "12 - VIDROS", "30 - CONJUNTO", "SALT", "N/A"],
                ["12180002", "12 - VIDROS", "10 - INSUMO", "SALT", "LATERAL"],
            ]
        )
        result = reconciliation.prepare_reconciliation(
            content,
            CATEGORY,
            _fields(),
            _rows(),
            lambda row, values: {"id": row["id"], "form_values": values},
        )
        self.assertEqual(result["total"], 2)
        self.assertEqual(
            result["payloads"][0]["form_values"][excel_bancos.PN_GROUP_FORM_KEY],
            ["10 - INSUMO"],
        )


if __name__ == "__main__":
    unittest.main()
