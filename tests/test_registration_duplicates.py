import unittest
from unittest.mock import patch

import excel_bancos
import supabase_store


class RegistrationDuplicateTests(unittest.TestCase):
    def setUp(self):
        self.existing = {
            "id": 2570,
            "sku": "30220074",
            "descricao_primaria": "CJ CAPA VESTIR BCO MOTORISTA LE COURVIN E/S/J CS PRETO/CINZA",
            "descricao_secundaria": "CJ CAPA VESTIR BCO MOTORISTA LE COURVIN E/S/J CS PRETO/CINZA COSTURA DIAMANTE LINHA CINZA",
            "sufixo": "2.2",
            "unidade": "cj",
        }

    def test_finds_duplicate_even_when_formatting_differs(self):
        duplicate = {
            **self.existing,
            "descricao_primaria": self.existing["descricao_primaria"].lower(),
            "descricao_secundaria": f"  {self.existing['descricao_secundaria']}  ",
        }
        with patch.object(supabase_store, "_request_all", return_value=[duplicate]):
            result = supabase_store._find_duplicate_registration(
                "cat_22_pecas_bco",
                self.existing["descricao_primaria"],
                self.existing["descricao_secundaria"],
                sufixo="2.2",
                unidade="CJ",
            )

        self.assertEqual(result["sku"], "30220074")

    def test_can_exclude_current_registration_when_editing(self):
        with patch.object(supabase_store, "_request_all", return_value=[self.existing]):
            result = supabase_store._find_duplicate_registration(
                "cat_22_pecas_bco",
                self.existing["descricao_primaria"],
                self.existing["descricao_secundaria"],
                sufixo="2.2",
                unidade="cj",
                exclude_id=2570,
            )

        self.assertIsNone(result)

    def test_inactive_duplicate_does_not_block_new_registration(self):
        inactive = {**self.existing, "ativo": False}

        def request_all(_table, query, limit=10000):
            del limit
            return [inactive] if ("ativo", "is.true") not in query else []

        with patch.object(supabase_store, "_request_all", side_effect=request_all):
            result = supabase_store._find_duplicate_registration(
                "cat_22_pecas_bco",
                self.existing["descricao_primaria"],
                self.existing["descricao_secundaria"],
                unidade="cj",
            )

        self.assertIsNone(result)

    def test_reactivation_checks_inactive_duplicates_too(self):
        inactive = {**self.existing, "ativo": False}

        def request_all(_table, query, limit=10000):
            del limit
            return [inactive] if ("ativo", "is.true") not in query else []

        with patch.object(supabase_store, "_request_all", side_effect=request_all):
            result = supabase_store._find_duplicate_registration(
                "cat_22_pecas_bco",
                self.existing["descricao_primaria"],
                self.existing["descricao_secundaria"],
                unidade="cj",
                include_inactive=True,
            )

        self.assertEqual(result["sku"], "30220074")

    def test_technical_suffix_difference_does_not_bypass_duplicate_check(self):
        existing = {**self.existing, "sufixo": ""}
        with patch.object(supabase_store, "_request_all", return_value=[existing]):
            result = supabase_store._find_duplicate_registration(
                "cat_22_pecas_bco",
                self.existing["descricao_primaria"],
                self.existing["descricao_secundaria"],
                sufixo="2.2",
                unidade="cj",
            )

        self.assertEqual(result["sku"], "30220074")

    def test_n_a_technical_option_does_not_bypass_duplicate_check(self):
        existing = {
            **self.existing,
            "descricao_primaria": "CJ CAPA VESTIR BCO MOTORISTA LE COURVIN E/S/J CS PRETO/CINZA COSTURA DIAMANTE LINHA CINZA",
            "sufixo": "",
            "field_values": {
                "cor": "7- PRETO/CINZA",
                "medida": "",
                "costura": "2- COSTURA DIAMANTE",
            },
        }
        with patch.object(supabase_store, "_request_all", return_value=[existing]):
            result = supabase_store._find_duplicate_registration(
                "cat_22_pecas_bco",
                "CJ CAPA VESTIR BCO MOTORISTA LE COURVIN E/S/J CS PRETO/CINZA",
                self.existing["descricao_secundaria"],
                sufixo="2.2",
                unidade="cj",
                field_values={
                    "cor": "7- PRETO/CINZA",
                    "medida": "3- N/A",
                    "costura": "2- COSTURA DIAMANTE",
                },
            )

        self.assertEqual(result["sku"], "30220074")

    def test_new_registration_is_rejected_before_insert(self):
        payload = {
            "category_key": "cat_22_pecas_bco",
            "category_label": "22 - PECAS BCO",
            "sku": "30220079",
            "descricao_primaria": self.existing["descricao_primaria"],
            "descricao_secundaria": self.existing["descricao_secundaria"],
            "sufixo": "2.2",
            "unidade": "cj",
        }
        with (
            patch.object(supabase_store, "_category", return_value={"key": "cat_22_pecas_bco"}),
            patch.object(excel_bancos, "get_banco_fields", return_value=[]),
            patch.object(supabase_store, "_next_sku", return_value="30220079"),
            patch.object(
                supabase_store,
                "_registration_payload",
                return_value=(payload, {"primaria": payload["descricao_primaria"], "secundaria": payload["descricao_secundaria"]}, False),
            ),
            patch.object(supabase_store, "_find_duplicate_registration", return_value=self.existing),
            patch.object(supabase_store, "_request") as request,
        ):
            with self.assertRaisesRegex(supabase_store.SupabaseStoreError, "30220074"):
                supabase_store.save_registration({"categoria": "cat_22_pecas_bco"})

        request.assert_not_called()

    def test_inactivation_is_allowed_when_another_active_duplicate_exists(self):
        current = {**self.existing, "ativo": True, "form_values": {}}
        payload = {
            **self.existing,
            "ativo": False,
            "category_key": "cat_22_pecas_bco",
            "field_values": {},
            "form_values": {},
        }
        updated = {**current, **payload}
        with (
            patch.object(supabase_store, "get_registration", return_value=current),
            patch.object(supabase_store, "_category", return_value={"key": "cat_22_pecas_bco", "label": "22 - PECAS BCO"}),
            patch.object(excel_bancos, "get_banco_fields", return_value=[]),
            patch.object(supabase_store, "_registration_structure_changed", return_value=False),
            patch.object(
                supabase_store,
                "_registration_payload",
                return_value=(payload, {"primaria": payload["descricao_primaria"], "secundaria": payload["descricao_secundaria"]}, False),
            ),
            patch.object(supabase_store, "_find_duplicate_registration") as find_duplicate,
            patch.object(supabase_store, "_request", return_value=[updated]),
            patch.object(supabase_store, "_set_catalog_bom_preference"),
            patch.object(supabase_store, "_record_audit_event"),
        ):
            result = supabase_store.update_registration(
                2571,
                {"categoria": "cat_22_pecas_bco", excel_bancos.PN_GROUP_FORM_KEY: "10"},
                actor="PAULO",
            )

        find_duplicate.assert_not_called()
        self.assertFalse(result["ativo"])


if __name__ == "__main__":
    unittest.main()
