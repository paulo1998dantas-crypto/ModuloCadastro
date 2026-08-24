import unittest
from unittest.mock import patch

import excel_bancos


class RevestimentoDescriptionRulesTests(unittest.TestCase):
    def test_na_de_revestimento_nao_compõe_descricao(self):
        fields = [
            {
                "key": "estagio",
                "label": "ESTAGIO",
                "scope": "primaria",
                "selection_mode": "unitaria",
                "description_order": 1,
            },
            {
                "key": "fornecedor",
                "label": "FORNECEDOR",
                "scope": "secundaria",
                "selection_mode": "unitaria",
                "description_order": 2,
            },
        ]

        description = excel_bancos.build_descriptions(
            fields,
            {"estagio": "1- CJ", "fornecedor": "2- N/A"},
            "cat_18_revestimento",
        )

        self.assertEqual(description["primaria"], "CJ")
        self.assertEqual(description["secundaria"], "CJ")
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
        self.assertEqual(payload["category_key"], "teste")
        self.assertEqual(payload["category_label"], "10 - TESTE")
        self.assertEqual(payload["sku"], "10100001")
        self.assertEqual(payload["unidade"], "pc")
        self.assertTrue(payload["ativo"])
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

    def test_refresh_applies_current_order_option_labels_and_conditional_rules(self):
        fields = [
            {
                "key": "gatilho",
                "label": "GATILHO",
                "scope": "primaria",
                "selection_mode": "unitaria",
                "description_order": 2,
                "options": ["1- ATIVO"],
            },
            {
                "key": "detalhe",
                "label": "DETALHE",
                "scope": "primaria",
                "selection_mode": "unitaria",
                "description_order": 1,
                "options": ["1- DETALHE NOVO"],
            },
            {
                "key": "promocao",
                "label": "PROMOCAO",
                "scope": "secundaria",
                "selection_mode": "unitaria",
                "description_order": 1,
                "options": ["1- PROMOVIDO"],
            },
            {
                "key": "oculto",
                "label": "OCULTO",
                "scope": "secundaria",
                "selection_mode": "unitaria",
                "description_order": 2,
                "options": ["1- NAO DEVE APARECER"],
            },
        ]
        row = {
            "id": "cadastro-1",
            "sku": "10100003",
            "unidade": "pc",
            "ativo": True,
            "form_values": {
                "gatilho": ["1- ATIVO"],
                "detalhe": ["1- DETALHE ANTIGO"],
                "promocao": ["1- PROMOVIDO"],
                "oculto": ["1- NAO DEVE APARECER"],
            },
        }
        rules = [
            {
                "key": "promover",
                "action": "set_primary",
                "source_type": "field",
                "source_field_key": "gatilho",
                "source_values": ["1- ATIVO"],
                "target_field_key": "promocao",
            },
            {
                "key": "ocultar",
                "action": "hide",
                "source_type": "field",
                "source_field_key": "gatilho",
                "source_values": ["1- ATIVO"],
                "target_field_key": "oculto",
            },
        ]

        with (
            patch.object(excel_bancos, "get_description_rules", return_value=[]),
            patch.object(excel_bancos, "get_conditional_rules", return_value=rules),
        ):
            payload = supabase_store._description_refresh_payload(row, self.category, fields)

        self.assertIsNotNone(payload)
        self.assertEqual(payload["descricao_primaria"], "DETALHE NOVO ATIVO PROMOVIDO")
        self.assertEqual(payload["descricao_secundaria"], "DETALHE NOVO ATIVO PROMOVIDO")
        self.assertEqual(payload["form_values"]["detalhe"], ["1- DETALHE NOVO"])
        self.assertNotIn("NAO DEVE APARECER", payload["descricao_secundaria"])


if __name__ == "__main__":
    unittest.main()
