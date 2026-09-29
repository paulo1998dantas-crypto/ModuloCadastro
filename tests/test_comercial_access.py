import unittest
from io import BytesIO
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import load_workbook

import main
import supabase_suprimentos


class ComercialAccessTests(unittest.TestCase):
    def test_comercial_allowlist_only_exposes_catalog_and_people_management(self):
        allowed = [
            ("GET", "/cadastros"),
            ("GET", "/cadastros/auditoria"),
            ("GET", "/cadastros/exportar"),
            ("GET", "/suprimentos"),
            ("GET", "/suprimentos/pessoas/template"),
            ("GET", "/suprimentos/pessoas/exportar"),
            ("POST", "/suprimentos/pessoas"),
            ("POST", "/suprimentos/pessoas/excluir"),
            ("POST", "/suprimentos/pessoas/upload"),
        ]
        denied = [
            ("GET", "/cadastros/42/editar"),
            ("GET", "/cadastro/bancos"),
            ("GET", "/bom"),
            ("GET", "/opcoes"),
            ("GET", "/layouts"),
            ("GET", "/suprimentos/regras/template"),
            ("POST", "/suprimentos/regras"),
            ("POST", "/suprimentos/relacoes"),
            ("PUT", "/suprimentos/pessoas"),
            ("DELETE", "/suprimentos/pessoas/excluir"),
        ]

        for method, path in allowed:
            with self.subTest(method=method, path=path):
                self.assertTrue(main._commercial_route_allowed(method, path))
        for method, path in denied:
            with self.subTest(method=method, path=path):
                self.assertFalse(main._commercial_route_allowed(method, path))

    def test_legacy_auth_allows_comercial_profile(self):
        self.assertTrue(
            main._legacy_cadastro_access_allowed(
                {"role": "COMERCIAL", "active": True}
            )
        )

    def test_comercial_works_as_an_additional_shared_profile_during_legacy_cutover(self):
        request = SimpleNamespace(state=SimpleNamespace())
        user = {"id": 92, "username": "vendas", "role": "PCP", "active": True}
        effective_access = {
            "active": True,
            "roles": ["PCP", "COMERCIAL"],
            "permissions": ["cadastro.access"],
        }
        with patch.object(main, "_read_session_payload", return_value={"u": "vendas"}), \
            patch.object(main, "_shared_user_lookup", return_value=user), \
            patch.object(main, "_shared_access_lookup", return_value=effective_access):
            self.assertTrue(main._legacy_cadastro_access_allowed({**user, "role": "OPERADOR"}))
            self.assertTrue(main._cadastro_is_comercial(request))

    def test_comercial_suprimentos_view_hides_other_maintenance_tables(self):
        html = main.templates.env.get_template("suprimentos.html").render(
            pessoas=[],
            pessoa_edicao={},
            processos={},
            regras=[],
            relacoes={},
            sucesso="",
            erro="",
            is_comercial=True,
            can_manage_people=True,
        )

        self.assertIn("Cadastro Geral de Pessoa", html)
        self.assertIn("Pessoas cadastradas", html)
        self.assertIn("Exportar Excel", html)
        self.assertIn('href="/cadastros/auditoria"', html)
        self.assertIn("Excluir", html)
        self.assertNotIn("Código identificador único", html)
        self.assertNotIn("Código do município", html)
        self.assertNotIn("Código da UF", html)
        self.assertNotIn("Código do país", html)
        self.assertNotIn("Limite de crédito", html)
        self.assertNotIn("Periodicidade compra/venda (dias)", html)
        self.assertNotIn("Valor mínimo de compra", html)
        self.assertNotIn(">Validação", html)
        self.assertNotIn("Processos Padrao", html)
        self.assertNotIn("Parametros de Item Relacionado", html)
        self.assertNotIn("Relacao Processo x Item", html)
        self.assertNotIn('href="/bom"', html)

    def test_people_export_contains_all_fields_and_keeps_formula_like_text_safe(self):
        person = {
            "id": 17,
            "nome_fantasia": "=HYPERLINK(\"https://example.invalid\")",
            "cnpj_cpf": "12.345.678/0001-90",
            "cliente": True,
        }
        with patch.object(supabase_suprimentos, "_all_rows", return_value=[person]):
            content = supabase_suprimentos.export_pessoas_xlsx()

        workbook = load_workbook(BytesIO(content), data_only=False)
        sheet = workbook.active
        headers = [cell.value for cell in sheet[1]]
        row = [cell.value for cell in sheet[2]]
        self.assertEqual("Pessoas cadastradas", sheet.title)
        self.assertEqual("id", headers[0])
        self.assertIn("nome_fantasia", headers)
        self.assertIn("data_nascimento_fundacao", headers)
        self.assertEqual(17, row[0])
        name_cell = sheet.cell(2, headers.index("nome_fantasia") + 1)
        self.assertEqual(person["nome_fantasia"], name_cell.value)
        self.assertEqual("s", name_cell.data_type)
        self.assertEqual("A2", sheet.freeze_panes)
        self.assertTrue(sheet.auto_filter.ref)
        workbook.close()


if __name__ == "__main__":
    unittest.main()
