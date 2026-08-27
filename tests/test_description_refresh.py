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
                "options": ["1- CJ"],
            },
            {
                "key": "fornecedor",
                "label": "FORNECEDOR",
                "scope": "secundaria",
                "selection_mode": "unitaria",
                "description_order": 2,
                "options": ["2- N/A"],
            },
        ]

        description = excel_bancos.build_descriptions(
            fields,
            {"estagio": "1- CJ", "fornecedor": "2- N/A"},
            "cat_18_revestimento",
        )

        self.assertEqual(description["primaria"], "CJ")
        self.assertEqual(description["secundaria"], "CJ")

    def test_na_e_omitido_em_qualquer_categoria_e_escopo(self):
        fields = [
            {
                "key": "modelo",
                "label": "MODELO",
                "scope": "primaria",
                "selection_mode": "unitaria",
                "description_order": 1,
                "options": ["1- N/A", "2- MODELO A"],
            },
            {
                "key": "acabamento",
                "label": "ACABAMENTO",
                "scope": "secundaria",
                "selection_mode": "unitaria",
                "description_order": 2,
                "options": ["1- N/A", "2- ACABAMENTO A"],
            },
        ]

        description = excel_bancos.build_descriptions(
            fields,
            {"modelo": "1- N/A", "acabamento": "1- N/A"},
            "cat_futura",
        )

        self.assertEqual(description["primaria"], "")
        self.assertEqual(description["secundaria"], "")
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

    def test_refresh_removes_orphan_catalog_values_and_preserves_free_text(self):
        fields = [
            {
                "key": "cor",
                "label": "COR",
                "scope": "primaria",
                "selection_mode": "multipla",
                "description_order": 1,
                "options": ["1- AZUL", "3- PRETO"],
            },
            {
                "key": "observacao",
                "label": "OBSERVACAO",
                "scope": "secundaria",
                "selection_mode": "unitaria",
                "description_order": 2,
                "free_text": True,
                "options": [],
            },
        ]
        rows = [
            {
                "id": "cadastro-1",
                "sku": "10100004",
                "unidade": "pc",
                "ativo": True,
                "form_values": {
                    "cor": ["1- AZUL", "2- VERDE EXCLUIDO"],
                    "observacao": ["TEXTO OPERACIONAL LIVRE"],
                    "possui_bom": True,
                },
            }
        ]
        calls = []

        def request(method, table, query=None, payload=None, prefer=""):
            calls.append((method, table, query, payload, prefer))
            return []

        with (
            patch.object(supabase_store, "_category", return_value=self.category),
            patch.object(excel_bancos, "get_banco_fields", return_value=fields),
            patch.object(supabase_store, "_request_all", return_value=rows),
            patch.object(supabase_store, "_request", side_effect=request),
        ):
            result = supabase_store.refresh_registration_descriptions("teste")

        registration_call = next(call for call in calls if call[1] == supabase_store.REGISTRATIONS_TABLE)
        payload = registration_call[3][0]
        self.assertEqual(payload["form_values"]["cor"], ["1- AZUL"])
        self.assertEqual(payload["form_values"]["observacao"], ["TEXTO OPERACIONAL LIVRE"])
        self.assertTrue(payload["form_values"]["possui_bom"])
        self.assertEqual(payload["descricao_primaria"], "AZUL")
        self.assertNotIn("VERDE", payload["descricao_secundaria"])
        self.assertEqual(result["removed_values"], 1)
        self.assertEqual(result["affected"], 1)
        self.assertEqual(result["removed_by_field"], {"cor": 1})

    def test_refresh_canonicalizes_by_unique_current_option_code(self):
        kept, removed, canonicalized = supabase_store._catalog_compatible_values(
            ["2- NOME ANTIGO"],
            ["1- AZUL", "2- NOME ATUAL"],
        )

        self.assertEqual(kept, ["2- NOME ATUAL"])
        self.assertEqual(removed, [])
        self.assertTrue(canonicalized)

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


class DeletedOptionReconciliationTests(unittest.TestCase):
    def setUp(self):
        self.category = {"key": "teste", "label": "10 - TESTE"}
        self.fields = [
            {
                "key": "cor",
                "label": "COR",
                "scope": "primaria",
                "selection_mode": "multipla",
                "description_order": 1,
                "options": ["1- AZUL", "3- PRETO"],
            }
        ]

    def test_deleted_option_is_removed_and_all_descriptions_are_recalculated(self):
        rows = [
            {
                "id": "cadastro-1",
                "sku": "10100001",
                "unidade": "pc",
                "ativo": True,
                "form_values": {
                    "cor": ["1- AZUL", "2- VERDE"],
                    "possui_bom": True,
                    "marcador_legado": {"origem": "importacao"},
                },
            },
            {
                "id": "cadastro-2",
                "sku": "10100002",
                "unidade": "pc",
                "ativo": False,
                "form_values": {"cor": ["3- PRETO"]},
            },
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
            result = supabase_store.reconcile_deleted_field_options(
                "teste",
                "cor",
                ["2- VERDE"],
            )

        registration_call = next(call for call in calls if call[1] == supabase_store.REGISTRATIONS_TABLE)
        payloads = registration_call[3]
        first_payload = next(payload for payload in payloads if payload["sku"] == "10100001")
        second_payload = next(payload for payload in payloads if payload["sku"] == "10100002")
        self.assertEqual(first_payload["form_values"]["cor"], ["1- AZUL"])
        self.assertEqual(first_payload["descricao_primaria"], "AZUL")
        self.assertTrue(first_payload["form_values"]["possui_bom"])
        self.assertEqual(first_payload["form_values"]["marcador_legado"], {"origem": "importacao"})
        self.assertEqual(second_payload["descricao_primaria"], "PRETO")
        self.assertFalse(second_payload["ativo"])
        self.assertEqual(result["affected"], 1)
        self.assertEqual(result["removed_values"], 1)
        self.assertEqual(result["recalculated"], 2)

    def test_deleted_code_removes_legacy_label_but_preserves_unrelated_unknown_value(self):
        kept, removed = supabase_store._values_without_deleted_options(
            ["2- VERDE LEGADO", "99- INFORMACAO HISTORICA"],
            ["1- AZUL", "3- PRETO"],
            ["2- VERDE"],
        )

        self.assertEqual(kept, ["99- INFORMACAO HISTORICA"])
        self.assertEqual(removed, ["2- VERDE LEGADO"])

    def test_remaining_option_with_same_code_is_canonicalized_not_removed(self):
        kept, removed = supabase_store._values_without_deleted_options(
            ["2- VERDE ANTIGO"],
            ["2- VERDE NOVO"],
            ["2- VERDE ANTIGO"],
        )

        self.assertEqual(kept, ["2- VERDE NOVO"])
        self.assertEqual(removed, [])


if __name__ == "__main__":
    unittest.main()
