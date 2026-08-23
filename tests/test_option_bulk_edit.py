import unittest
from unittest.mock import patch

import excel_bancos


def _catalog():
    return {
        "active_category": "acessorios",
        "categories": [
            {
                "key": "acessorios",
                "label": "ACESSORIOS",
                "fields": [
                    {
                        "key": "cor",
                        "label": "COR",
                        "scope": "primaria",
                        "selection_mode": "unitaria",
                        "options": ["1- AZUL", "2- VERDE", "3- PRETO"],
                    }
                ],
                "conditional_rules": [],
            }
        ],
    }


class OptionBulkEditTests(unittest.TestCase):
    def test_updates_and_deletes_in_one_catalog_save(self):
        catalog = _catalog()
        with patch.object(excel_bancos, "load_catalog", return_value=catalog), patch.object(
            excel_bancos, "save_catalog"
        ) as save_catalog:
            result = excel_bancos.update_field_options(
                "acessorios",
                "cor",
                [1, 2, 3],
                ["AZUL MARINHO", "VERDE", "PRETO"],
                delete_row_values=[2],
            )

        self.assertEqual(catalog["categories"][0]["fields"][0]["options"], ["1- AZUL MARINHO", "3- PRETO"])
        self.assertEqual(result["deleted"], ["2- VERDE"])
        self.assertEqual(result["deleted_count"], 1)
        self.assertEqual(result["option_rows"], [{"row": 1, "value": "1- AZUL MARINHO"}, {"row": 2, "value": "3- PRETO"}])
        save_catalog.assert_called_once_with(catalog)

    def test_rejects_duplicate_against_a_remaining_option_atomically(self):
        catalog = _catalog()
        original = list(catalog["categories"][0]["fields"][0]["options"])
        with patch.object(excel_bancos, "load_catalog", return_value=catalog), patch.object(
            excel_bancos, "save_catalog"
        ) as save_catalog:
            with self.assertRaisesRegex(ValueError, "existe"):
                excel_bancos.update_field_options(
                    "acessorios",
                    "cor",
                    [1, 2, 3],
                    ["PRETO", "VERDE", "PRETO"],
                    delete_row_values=[2],
                )

        self.assertEqual(catalog["categories"][0]["fields"][0]["options"], original)
        save_catalog.assert_not_called()

    def test_adds_multiple_options_with_one_persistence(self):
        catalog = _catalog()
        with patch.object(excel_bancos, "load_catalog", return_value=catalog), patch.object(
            excel_bancos, "save_catalog"
        ) as save_catalog:
            result = excel_bancos.add_field_options("acessorios", "cor", ["BRANCO", "CINZA"])

        self.assertEqual(result["count"], 2)
        self.assertEqual(result["options"], ["4- BRANCO", "5- CINZA"])
        self.assertEqual(result["rows"], [4, 5])
        save_catalog.assert_called_once_with(catalog)


if __name__ == "__main__":
    unittest.main()
