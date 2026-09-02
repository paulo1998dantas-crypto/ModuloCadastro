import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

import supabase_store


class BomInactiveFilterTests(unittest.TestCase):
    def setUp(self):
        self.headers = [
            {"id": 1, "parent_sku": "30180001", "parent_descricao": "ATIVO"},
            {"id": 2, "parent_sku": "30180002", "parent_descricao": "INATIVO"},
        ]
        self.components = [
            {"id": 11, "bom_id": 1, "component_sku": "10180001", "quantidade": 1, "ordem": 1},
            {"id": 12, "bom_id": 2, "component_sku": "10180001", "quantidade": 1, "ordem": 1},
        ]
        self.catalog = {
            "30180001": {"descricao_primaria": "PAI ATIVO", "unidade": "cj", "ativo": True},
            "30180002": {"descricao_primaria": "PAI INATIVO", "unidade": "cj", "ativo": False},
            "10180001": {"descricao_primaria": "COMPONENTE", "unidade": "pc", "ativo": True},
        }

    def _request_all(self, table, *args, **kwargs):
        if table == supabase_store.BOM_HEADERS_TABLE:
            return self.headers
        if table == supabase_store.BOM_COMPONENTS_TABLE:
            return self.components
        return []

    def test_inactive_parent_is_hidden_by_default(self):
        with (
            patch.object(supabase_store, "_request_all", side_effect=self._request_all),
            patch.object(supabase_store, "_catalog_data_by_sku", return_value=self.catalog),
        ):
            rows = supabase_store.list_boms()

        self.assertEqual([row["parent_sku"] for row in rows], ["30180001"])
        self.assertEqual(rows[0]["parent_status"], "ATIVO")

    def test_inactive_parent_is_returned_only_when_requested(self):
        with (
            patch.object(supabase_store, "_request_all", side_effect=self._request_all),
            patch.object(supabase_store, "_catalog_data_by_sku", return_value=self.catalog),
        ):
            rows = supabase_store.list_boms(include_inactive=True)

        self.assertEqual(len(rows), 2)
        inactive = next(row for row in rows if row["parent_sku"] == "30180002")
        self.assertFalse(inactive["parent_active"])
        self.assertEqual(inactive["parent_status"], "INATIVO")

    def test_export_includes_parent_status_and_respects_flag(self):
        rows = [
            {
                "parent_category_label": "18 - REVESTIMENTO",
                "parent_sku": "30180002",
                "display_parent_sku": "30180002",
                "parent_descricao": "PAI INATIVO",
                "parent_active": False,
                "parent_status": "INATIVO",
                "components": [
                    {
                        "component_sku": "10180001",
                        "display_component_sku": "10180001",
                        "component_descricao": "COMPONENTE",
                        "unidade": "pc",
                        "quantidade": 1,
                    }
                ],
            }
        ]
        with tempfile.TemporaryDirectory() as temp_dir:
            with (
                patch.object(supabase_store, "EXPORT_DIR", Path(temp_dir)),
                patch.object(supabase_store, "list_boms", return_value=rows) as list_boms,
            ):
                output = supabase_store.export_boms(include_inactive=True)

            list_boms.assert_called_once_with(
                category_key="",
                parent_query="",
                component_query="",
                include_inactive=True,
                limit=5000,
            )
            workbook = load_workbook(output, data_only=True)
            try:
                sheet = workbook["BOM"]
                self.assertEqual(sheet["D1"].value, "status_item_pai")
                self.assertEqual(sheet["D2"].value, "INATIVO")
            finally:
                workbook.close()


if __name__ == "__main__":
    unittest.main()
