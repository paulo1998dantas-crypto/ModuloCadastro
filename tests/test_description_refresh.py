import unittest
from unittest.mock import patch

import excel_bancos
import supabase_store


class DescriptionRefreshTests(unittest.TestCase):
    def setUp(self):
        self.category = {"key": "teste", "label": "10 - TESTE"}
        self.fields = [
            {
                "key": "modelo",
                "label": "MODELO",
                "scope": "primaria",
                "selection_mode": "unitaria",
                "description_order": 1,
                "options": ["1- MODELO A"],
            }
        ]

    def test_refresh_rebuilds_descriptions_without_changing_operational_data(self):
        rows = [
            {
                "id": "2e5de3a1-3d15-4f92-a7d8-8d9822e641e7",
                "sku": "10100001",
                "unidade": "pc",
                "form_values": {
                    "modelo": ["1- MODELO A"],
                    "possui_bom": True,
                    "marcador_legado": {"origem": "importacao"},
                },
            }
        ]
        calls = []

        def request(method, table, query=None, payload=None, prefer=""):
            calls.append((method, table, query, payload, prefer))
            return []

        with (
            patch.object(supabase_store, "_category", return_value=self.category),
            patch.object(excel_bancos, "get_banco_fields", return_value=self.fields),
            patch.object(supabase_store, "_request_all", return_value=rows),
            patch.object(supabase_store, "_request", side_effect=request),
        ):
            result = supabase_store.refresh_registration_descriptions("teste")

        self.assertEqual(result["updated"], 1)
        self.assertEqual(result["skipped"], 0)
        post = next(call for call in calls if call[0] == "POST")
        payload = post[3][0]
        self.assertEqual(payload["id"], rows[0]["id"])
        self.assertEqual(payload["descricao_primaria"], "MODELO A")
        self.assertTrue(payload["form_values"]["possui_bom"])
        self.assertEqual(payload["form_values"]["marcador_legado"], {"origem": "importacao"})
        self.assertEqual(post[4], "resolution=merge-duplicates,return=minimal")

    def test_refresh_skips_records_without_structured_field_values(self):
        rows = [{"id": "a", "sku": "10100002", "unidade": "pc", "form_values": None}]
        with (
            patch.object(supabase_store, "_category", return_value=self.category),
            patch.object(excel_bancos, "get_banco_fields", return_value=self.fields),
            patch.object(supabase_store, "_request_all", return_value=rows),
            patch.object(supabase_store, "_request") as request,
        ):
            result = supabase_store.refresh_registration_descriptions("teste")

        self.assertEqual(result["updated"], 0)
        self.assertEqual(result["skipped"], 1)
        self.assertEqual(result["skipped_identifiers"], ["10100002"])
        request.assert_not_called()


if __name__ == "__main__":
    unittest.main()
