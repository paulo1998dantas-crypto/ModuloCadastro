import unittest
from io import BytesIO
from unittest.mock import patch

from openpyxl import load_workbook

import supabase_suprimentos

class PessoaSupabaseTests(unittest.TestCase):
    def test_normalizar_pessoa_preserva_endereco_contato_e_parametros(self):
        pessoa = supabase_suprimentos.normalizar_pessoa(
            {
                "nome_fantasia": "Cliente Teste",
                "cnpj_cpf": "12.345.678/0001-90",
                "logradouro": "Rua das Flores",
                "logradouro_numero": "123",
                "complemento": "Sala 2",
                "bairro": "Centro",
                "cidade": "Sao Paulo",
                "uf": "SP",
                "cep": "01000-000",
                "telefone": "(11) 3333-4444",
                "whatsapp": "(11) 99999-8888",
                "cliente": "1",
                "limite_credito": "1.250,50",
                "data_registro": "2026-07-16",
            }
        )

        self.assertEqual(pessoa["logradouro"], "Rua das Flores")
        self.assertEqual(pessoa["logradouro_numero"], "123")
        self.assertTrue(pessoa["cliente"])
        self.assertEqual(pessoa["limite_credito"], 1250.5)
        self.assertEqual(pessoa["data_registro"], "2026-07-16")
        self.assertIn("Rua das Flores", pessoa["search_text"])
        self.assertIn("01000-000", pessoa["search_text"])

    def test_atualizar_pessoa_usa_patch_sem_criar_novo_registro(self):
        before = {
            "id": 42,
            "identificador": "CLIENTE-42",
            "nome_fantasia": "Cliente original",
            "payload": {},
        }
        after = {**before, "nome_fantasia": "Cliente Atualizado"}
        with (
            patch.object(supabase_suprimentos, "obter_pessoa", side_effect=[before, after]),
            patch.object(supabase_suprimentos, "_request") as request_mock,
            patch.object(supabase_suprimentos, "_registrar_auditoria_pessoa"),
        ):
            count = supabase_suprimentos.atualizar_pessoa(
                42,
                {
                    "identificador": "CLIENTE-42",
                    "nome_fantasia": "Cliente Atualizado",
                    "logradouro": "Avenida Brasil",
                    "cliente": True,
                },
            )

        self.assertEqual(count, 1)
        request_mock.assert_called_once()
        args, kwargs = request_mock.call_args
        self.assertEqual(args[:2], ("PATCH", supabase_suprimentos.PESSOAS_TABLE))
        self.assertIn(("id", "eq.42"), kwargs["query"])
        self.assertEqual(kwargs["payload"]["logradouro"], "Avenida Brasil")

    @patch("supabase_suprimentos._all_rows")
    def test_exportar_pessoas_gera_planilha_com_registros_e_cabecalho(self, all_rows):
        all_rows.return_value = [
            {"id": 8, "nome_fantasia": "Empresa Teste", "fornecedor": True}
        ]

        workbook = load_workbook(
            BytesIO(supabase_suprimentos.export_pessoas_xlsx()), data_only=True
        )
        sheet = workbook.active
        headers = [cell.value for cell in sheet[1]]
        self.assertEqual("id", headers[0])
        self.assertIn("fornecedor", headers)
        self.assertEqual("Empresa Teste", sheet.cell(2, headers.index("nome_fantasia") + 1).value)
        self.assertTrue(sheet.auto_filter.ref)
        workbook.close()

    @patch("supabase_suprimentos._all_rows")
    def test_exportar_pessoas_aplica_busca_da_tabela(self, all_rows):
        all_rows.return_value = [
            {"id": 8, "nome_fantasia": "Empresa Norte", "cidade": "Manaus"},
            {"id": 9, "nome_fantasia": "Empresa Sul", "cidade": "Curitiba"},
        ]

        workbook = load_workbook(
            BytesIO(supabase_suprimentos.export_pessoas_xlsx("Manaus")),
            data_only=True,
        )
        sheet = workbook.active
        self.assertEqual(2, sheet.max_row)
        self.assertEqual(
            "Empresa Norte",
            sheet.cell(2, [cell.value for cell in sheet[1]].index("nome_fantasia") + 1).value,
        )
        workbook.close()


if __name__ == "__main__":
    unittest.main()
