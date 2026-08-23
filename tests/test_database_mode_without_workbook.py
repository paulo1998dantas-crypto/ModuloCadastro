import os
import unittest
from unittest.mock import patch

import excel_bancos


class DatabaseModeWithoutWorkbookTests(unittest.TestCase):
    def test_supabase_mode_does_not_sync_legacy_workbook(self):
        with (
            patch.dict(os.environ, {"CADASTRO_SAVE_MODE": "supabase"}, clear=False),
            patch.object(excel_bancos, "active_workbook_path") as workbook_path,
            patch.object(excel_bancos, "sync_workbook_headers") as sync_headers,
        ):
            excel_bancos.sync_workbook_structure("cat_10_ar_condicionado")

        workbook_path.assert_not_called()
        sync_headers.assert_not_called()

    def test_local_mode_keeps_optional_workbook_compatibility(self):
        with (
            patch.dict(os.environ, {"CADASTRO_SAVE_MODE": "local"}, clear=False),
            patch.object(excel_bancos, "active_workbook_path", return_value="cadastro.xlsx") as workbook_path,
            patch.object(excel_bancos, "sync_workbook_headers") as sync_headers,
        ):
            excel_bancos.sync_workbook_structure("cat_10_ar_condicionado")

        workbook_path.assert_called_once_with()
        sync_headers.assert_called_once_with("cadastro.xlsx", "cat_10_ar_condicionado")


class CatalogFieldPersistenceTests(unittest.TestCase):
    def setUp(self):
        self.catalog = {
            "version": 2,
            "active_category": "cat_10_ar_condicionado",
            "categories": [
                {
                    "key": "cat_10_ar_condicionado",
                    "label": "10 - AR CONDICIONADO",
                    "sheet_name": "10 - AR CONDICIONADO",
                    "fields": [
                        {
                            "key": "descritor_base",
                            "label": "DESCRITOR BASE",
                            "scope": "primaria",
                            "selection_mode": "unitaria",
                            "description_order": 1,
                            "options": ["1- EVAPORADOR"],
                        }
                    ],
                    "conditional_rules": [],
                }
            ],
            "pn_groups": [],
        }

    def test_update_field_changes_persisted_catalog_object(self):
        with (
            patch.object(excel_bancos, "load_catalog", return_value=self.catalog),
            patch.object(excel_bancos, "save_catalog") as save_catalog,
            patch.object(excel_bancos, "sync_workbook_structure"),
        ):
            result = excel_bancos.update_field(
                "cat_10_ar_condicionado",
                "descritor_base",
                "ITEM",
                "primaria",
                "unitaria",
            )

        self.assertEqual("ITEM", result["field"])
        self.assertEqual("ITEM", self.catalog["categories"][0]["fields"][0]["label"])
        self.assertIs(save_catalog.call_args.args[0], self.catalog)

    def test_update_field_option_changes_persisted_catalog_object(self):
        with (
            patch.object(excel_bancos, "load_catalog", return_value=self.catalog),
            patch.object(excel_bancos, "save_catalog"),
        ):
            excel_bancos.update_field_option(
                "cat_10_ar_condicionado",
                "descritor_base",
                1,
                "CONDENSADOR",
            )

        self.assertEqual(
            ["1- CONDENSADOR"],
            self.catalog["categories"][0]["fields"][0]["options"],
        )


if __name__ == "__main__":
    unittest.main()
