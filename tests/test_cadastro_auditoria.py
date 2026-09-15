import unittest
from unittest.mock import patch

import supabase_store


class CadastroAuditoriaTests(unittest.TestCase):
    def test_audit_diff_keeps_only_changed_fields(self):
        before = {"sku": "30220074", "ativo": True, "form_values": {"cor": "PRETO"}}
        after = {"sku": "30220074", "ativo": False, "form_values": {"cor": "CINZA"}}

        changes = supabase_store._audit_diff(before, after)

        self.assertEqual(set(changes), {"ativo", "form_values"})
        self.assertEqual(changes["ativo"], {"antes": True, "depois": False})

    def test_list_audit_events_merges_catalog_and_parameter_history(self):
        catalog_events = [
            {
                "id": 10,
                "registration_id": 81,
                "sku": "30220074",
                "previous_sku": "",
                "category_key": "cat_22_pecas_bco",
                "category_label": "22 - PECAS BCO",
                "action": "alteracao",
                "actor": "PAULO",
                "summary": "Cadastro alterado.",
                "details": {"alteracoes": {"ativo": {"antes": True, "depois": False}}},
                "created_at": "2026-09-15T12:00:00+00:00",
            }
        ]
        parameter_events = [
            {
                "id": 11,
                "registration_id": 81,
                "sku": "30220074",
                "action": "UPDATE",
                "actor": "PAULO",
                "before_data": {"producao_dias": 2},
                "after_data": {"producao_dias": 3},
                "changed_at": "2026-09-15T13:00:00+00:00",
            }
        ]

        def request_all(table, params, limit=10000):
            if table == supabase_store.AUDIT_TABLE:
                return catalog_events
            self.assertEqual(table, "cadastro_item_parametros_historico")
            return parameter_events

        with patch.object(supabase_store, "_request_all", side_effect=request_all):
            result = supabase_store.list_audit_events(sku="30220074", actor="PAULO")

        self.assertEqual([row["action"] for row in result], ["parametro_alteracao", "alteracao"])
        self.assertEqual(result[0]["previous_sku"], "")
        self.assertEqual(result[0]["details"]["depois"], {"producao_dias": 3})

    def test_list_audit_events_can_filter_sku_migration_by_old_sku(self):
        event = {
            "id": 20,
            "registration_id": 81,
            "sku": "20140031",
            "previous_sku": "30140027",
            "category_key": "cat_14_piso",
            "category_label": "14 - PISO",
            "action": "migracao_sku",
            "actor": "PAULO",
            "summary": "SKU substituído.",
            "details": {},
            "created_at": "2026-09-15T14:00:00+00:00",
        }

        with patch.object(supabase_store, "_request_all", side_effect=lambda *args, **kwargs: [event] if args[0] == supabase_store.AUDIT_TABLE else []):
            result = supabase_store.list_audit_events(sku="30140027")

        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["action"], "migracao_sku")


if __name__ == "__main__":
    unittest.main()
