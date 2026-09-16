import unittest
from unittest.mock import patch

import main
import supabase_store


class CadastroAuditoriaTests(unittest.TestCase):
    def test_audit_change_rows_are_recursive_and_ignore_technical_metadata(self):
        event = {
            "details": {
                "antes": {
                    "sku": "30220074",
                    "updated_at": "2026-09-15T10:00:00+00:00",
                    "form_values": {"cor": "1- PRETO", "costura": "2- RETA"},
                },
                "depois": {
                    "sku": "30220074",
                    "updated_at": "2026-09-15T10:01:00+00:00",
                    "form_values": {"cor": "2- CINZA", "costura": "2- RETA"},
                },
            }
        }

        rows = main._audit_change_rows(event)

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["label"], "Campos do formulário › Cor")
        self.assertEqual(rows[0]["before"], "1- PRETO")
        self.assertEqual(rows[0]["after"], "2- CINZA")

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
            if table == supabase_store.REGISTRATIONS_TABLE:
                return []
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

    def test_new_registration_persists_creator_and_user_id(self):
        payload = {
            "category_key": "cat_22_pecas_bco",
            "category_label": "22 - PECAS BCO",
            "sku": "30220080",
            "descricao_primaria": "ITEM NOVO",
            "descricao_secundaria": "",
            "sufixo": "",
            "unidade": "pc",
            "field_values": {},
        }
        with (
            patch.object(supabase_store, "_category", return_value={"key": "cat_22_pecas_bco", "label": "22 - PECAS BCO"}),
            patch.object(supabase_store.excel_bancos, "get_banco_fields", return_value=[]),
            patch.object(supabase_store, "_next_sku", return_value="30220080"),
            patch.object(
                supabase_store,
                "_registration_payload",
                return_value=(payload, {"primaria": "ITEM NOVO", "secundaria": ""}, False),
            ),
            patch.object(supabase_store, "_find_duplicate_registration", return_value=None),
            patch.object(supabase_store, "_request", return_value=[{"id": 88, **payload}]) as request,
            patch.object(supabase_store, "_record_audit_event") as audit,
        ):
            result = supabase_store.save_registration(
                {"categoria": "cat_22_pecas_bco"},
                actor="PAULO",
                actor_user_id=9,
            )

        saved_payload = request.call_args.kwargs["payload"]
        self.assertEqual(saved_payload["created_by"], "PAULO")
        self.assertEqual(saved_payload["created_by_user_id"], 9)
        self.assertEqual(result["created_by"], "PAULO")
        audit.assert_called_once()
        self.assertEqual(audit.call_args.kwargs["actor_user_id"], 9)

    def test_audit_user_filter_also_matches_registration_creator(self):
        event = {
            "id": 12,
            "registration_id": 81,
            "sku": "30220074",
            "action": "criacao",
            "actor": "migracao",
            "summary": "Cadastro criado.",
            "details": {},
            "created_at": "2026-09-15T12:00:00+00:00",
        }

        def request_all(table, params, limit=10000):
            if table == supabase_store.REGISTRATIONS_TABLE:
                return [{"id": 81}]
            if table == supabase_store.AUDIT_TABLE:
                return [event]
            return []

        with patch.object(supabase_store, "_request_all", side_effect=request_all):
            result = supabase_store.list_audit_events(actor="PAULO")

        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["registration_id"], 81)


if __name__ == "__main__":
    unittest.main()
