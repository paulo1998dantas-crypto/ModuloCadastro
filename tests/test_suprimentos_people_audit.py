import unittest
from unittest.mock import patch

import supabase_suprimentos
import supabase_store


class SuprimentosPeopleAuditTests(unittest.TestCase):
    def test_new_person_gets_creator_and_brazil_creation_date_and_is_audited(self):
        saved = {
            "id": 25,
            "identificador": "12.345.678/0001-90",
            "nome_fantasia": "Fornecedor novo",
            "data_registro": "2026-09-29",
            "payload": {"identificador_cadastro": "PAULO"},
        }
        with (
            patch.object(supabase_suprimentos, "_all_rows", return_value=[]),
            patch.object(supabase_suprimentos, "_request", return_value=[saved]) as request,
            patch.object(supabase_suprimentos, "_registrar_auditoria_pessoa") as audit,
            patch.object(supabase_suprimentos, "data_registro_atual", return_value="2026-09-29"),
        ):
            count = supabase_suprimentos.salvar_pessoas(
                [{"nome_fantasia": "Fornecedor novo", "cnpj_cpf": "12.345.678/0001-90"}],
                actor="PAULO",
                actor_user_id=7,
            )

        self.assertEqual(count, 1)
        payload = request.call_args.kwargs["payload"][0]
        self.assertEqual(payload["data_registro"], "2026-09-29")
        self.assertEqual(payload["payload"]["identificador_cadastro"], "PAULO")
        self.assertEqual(payload["payload"]["identificador_cadastro_user_id"], 7)
        audit.assert_called_once()
        self.assertEqual(audit.call_args.args[0], "pessoa_criacao")
        self.assertEqual(audit.call_args.kwargs["actor_user_id"], 7)

    def test_person_edit_preserves_creator_date_and_legacy_hidden_fields(self):
        before = {
            "id": 31,
            "identificador": "LEGACY-31",
            "nome_fantasia": "Nome antigo",
            "data_registro": "2024-02-03",
            "codigo_municipio": "3550308",
            "limite_credito": 500,
            "payload": {"identificador_cadastro": "CRIADOR"},
        }
        after = {**before, "nome_fantasia": "Nome atualizado"}
        with (
            patch.object(supabase_suprimentos, "obter_pessoa", side_effect=[before, after]),
            patch.object(supabase_suprimentos, "_request") as request,
            patch.object(supabase_suprimentos, "_registrar_auditoria_pessoa") as audit,
        ):
            count = supabase_suprimentos.atualizar_pessoa(
                31,
                {"nome_fantasia": "Nome atualizado"},
                actor="EDITOR",
                actor_user_id=8,
            )

        self.assertEqual(count, 1)
        payload = request.call_args.kwargs["payload"]
        self.assertEqual(payload["data_registro"], "2024-02-03")
        self.assertEqual(payload["codigo_municipio"], "3550308")
        self.assertEqual(payload["limite_credito"], 500)
        self.assertEqual(payload["payload"]["identificador_cadastro"], "CRIADOR")
        audit.assert_called_once()
        self.assertEqual(audit.call_args.args[0], "pessoa_alteracao")
        self.assertEqual(audit.call_args.kwargs["actor"], "EDITOR")

    def test_person_delete_records_before_snapshot(self):
        before = {"id": 41, "identificador": "P-41", "nome_fantasia": "Excluir"}
        with (
            patch.object(supabase_suprimentos, "obter_pessoa", return_value=before),
            patch.object(supabase_suprimentos, "_request", return_value=[before]),
            patch.object(supabase_suprimentos, "_registrar_auditoria_pessoa") as audit,
        ):
            count = supabase_suprimentos.excluir_pessoa(41, actor="PAULO", actor_user_id=7)

        self.assertEqual(count, 1)
        audit.assert_called_once()
        self.assertEqual(audit.call_args.args[0], "pessoa_exclusao")
        self.assertEqual(audit.call_args.kwargs["before"], before)
        self.assertIsNone(audit.call_args.kwargs["after"])

    def test_people_audit_can_be_found_by_name_or_document(self):
        event = {
            "id": 73,
            "registration_id": 51,
            "sku": "PESSOA-51",
            "previous_sku": "",
            "category_key": "suprimentos_pessoas",
            "category_label": "Pessoas (Suprimentos)",
            "action": "pessoa_alteracao",
            "actor": "PAULO",
            "summary": "Pessoa alterada: Acme Ltda.",
            "details": {"antes": {"cnpj_cpf": "12345678000190"}, "depois": {}},
            "created_at": "2026-09-29T12:00:00+00:00",
        }
        with patch.object(
            supabase_store,
            "_request_all",
            side_effect=lambda table, *args, **kwargs: [event]
            if table == supabase_store.AUDIT_TABLE
            else [],
        ):
            by_name = supabase_store.list_audit_events(sku="Acme")
            by_document = supabase_store.list_audit_events(sku="12345678")

        self.assertEqual([event["id"] for event in by_name], [73])
        self.assertEqual([event["id"] for event in by_document], [73])


if __name__ == "__main__":
    unittest.main()
