import asyncio
import unittest
from unittest.mock import patch

import main
import supabase_suprimentos


class SuprimentosRuleEditTests(unittest.TestCase):
    def test_update_rule_patches_only_selected_rule_and_normalizes_options(self):
        with patch.object(supabase_suprimentos, "_request", return_value=[{"rule_id": "regra-7"}]) as request:
            count = supabase_suprimentos.atualizar_regra(
                "regra-7",
                {
                    "gatilho": "SKU-PAI",
                    "opcoes": ["SKU-A", " SKU-A ", "SKU-B"],
                    "quantidade": 2.5,
                    "quantidade_editavel": True,
                },
            )

        self.assertEqual(count, 1)
        args = request.call_args
        self.assertEqual(args.args[:2], ("PATCH", supabase_suprimentos.REGRAS_TABLE))
        self.assertEqual(args.kwargs["query"], [("rule_id", "eq.regra-7")])
        self.assertEqual(args.kwargs["payload"], {
            "gatilho": "SKU-PAI",
            "opcoes": ["SKU-A", "SKU-B"],
            "quantidade": 2.5,
            "quantidade_editavel": True,
        })

    def test_update_rule_rejects_nonpositive_quantity_and_missing_row(self):
        with self.assertRaisesRegex(ValueError, "maior que zero"):
            supabase_suprimentos.atualizar_regra(
                "regra-1", {"gatilho": "PAI", "opcoes": ["FILHO"], "quantidade": 0}
            )
        with patch.object(supabase_suprimentos, "_request", return_value=[]):
            with self.assertRaisesRegex(ValueError, "não encontrada"):
                supabase_suprimentos.atualizar_regra(
                    "regra-404", {"gatilho": "PAI", "opcoes": ["FILHO"], "quantidade": 1}
                )

    def test_update_item_process_relation_can_change_key_with_one_patch(self):
        with patch.object(supabase_suprimentos, "_request", return_value=[{"item_codigo": "NOVO"}]) as request:
            count = supabase_suprimentos.atualizar_relacao(
                "ANTIGO", "NOVO", ["Preparação", "Preparação", "Elétrica"]
            )

        self.assertEqual(count, 1)
        args = request.call_args
        self.assertEqual(args.args[:2], ("PATCH", supabase_suprimentos.RELACOES_TABLE))
        self.assertEqual(args.kwargs["query"], [("item_codigo", "eq.ANTIGO")])
        self.assertEqual(args.kwargs["payload"], {
            "item_codigo": "NOVO",
            "processos": ["Preparação", "Elétrica"],
        })

    def test_edit_routes_call_targeted_update_and_redirect_on_success(self):
        with patch.object(supabase_suprimentos, "atualizar_regra", return_value=1) as update_rule:
            response = asyncio.run(main.suprimentos_regras_editar(
                rule_id="regra-2",
                gatilho=" PAI ",
                opcoes="FILHO A; FILHO B",
                quantidade="2,5",
                quantidade_editavel="1",
            ))
        self.assertEqual(response.status_code, 303)
        update_rule.assert_called_once_with("regra-2", {
            "gatilho": "PAI",
            "opcoes": ["FILHO A", "FILHO B"],
            "quantidade": 2.5,
            "quantidade_editavel": True,
        })

        with patch.object(supabase_suprimentos, "atualizar_relacao", return_value=1) as update_relation:
            response = asyncio.run(main.suprimentos_relacoes_editar(
                item_codigo_anterior="001",
                item_codigo="002",
                processos="Preparação; Elétrica",
            ))
        self.assertEqual(response.status_code, 303)
        update_relation.assert_called_once_with("001", "002", ["Preparação", "Elétrica"])

    def test_template_exposes_edit_links_for_both_rule_tables(self):
        html = main.templates.env.get_template("suprimentos.html").render(
            pessoas=[], pessoa_edicao={}, pessoa_criador="ADMIN", data_registro_atual="2026-10-27",
            processos={}, regras=[{
                "id": "regra-9", "gatilho": "PAI", "opcoes": ["FILHO"],
                "quantidade": 2, "quantidade_editavel": True,
            }], relacoes={"001": ["Preparação"]}, regra_edicao={}, relacao_edicao={},
            sucesso="", erro="", is_comercial=False, can_manage_people=True,
        )

        self.assertIn("href=\"/suprimentos?editar_regra=regra-9#parametros-item-relacionado\"", html)
        self.assertIn("href=\"/suprimentos?editar_relacao=001#relacao-processo-item\"", html)

        edit_html = main.templates.env.get_template("suprimentos.html").render(
            pessoas=[], pessoa_edicao={}, pessoa_criador="ADMIN", data_registro_atual="2026-10-27",
            processos={}, regras=[], relacoes={}, regra_edicao={
                "id": "regra-9", "gatilho": "PAI", "opcoes": ["FILHO"],
                "quantidade": 2, "quantidade_editavel": True,
            }, relacao_edicao={"item_codigo": "001", "processos": ["Preparação"]},
            sucesso="", erro="", is_comercial=False, can_manage_people=True,
        )
        self.assertIn('action="/suprimentos/regras/editar"', edit_html)
        self.assertIn('action="/suprimentos/relacoes/editar"', edit_html)
        self.assertIn('name="item_codigo_anterior" value="001"', edit_html)


if __name__ == "__main__":
    unittest.main()
