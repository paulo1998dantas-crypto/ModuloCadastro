import unittest
from unittest.mock import patch

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

    @patch("supabase_suprimentos._request")
    def test_atualizar_pessoa_usa_patch_sem_criar_novo_registro(self, request_mock):
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


if __name__ == "__main__":
    unittest.main()
