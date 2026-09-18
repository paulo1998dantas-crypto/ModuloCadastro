import unittest
from unittest.mock import patch
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import supabase_store


class EquivalenceStoreTests(unittest.TestCase):
    def test_group_view_normalizes_factor_and_shows_active_registration(self):
        group = {"id": "group-1", "codigo": "EQ-CAPA", "nome": "Capa equivalentes"}
        members = [
            {"sku": "10220077", "prioridade": 20, "fator_unidade_funcional": "2"},
            {"sku": "10220078", "prioridade": 10, "fator_unidade_funcional": "1"},
        ]
        catalog = {
            "10220077": {"descricao_primaria": "Capa versão dois", "unidade": "pc", "ativo": True},
            "10220078": {"descricao_primaria": "Capa versão um", "unidade": "pc", "ativo": True},
        }

        result = supabase_store._equivalence_group_view(group, members, catalog)

        self.assertEqual(["10220078", "10220077"], [member["sku"] for member in result["members"]])
        self.assertEqual(1.0, result["members"][0]["fator_unidade_funcional"])
        self.assertEqual(2.0, result["members"][1]["fator_unidade_funcional"])
        self.assertTrue(result["members"][0]["registration_active"])

    def test_member_requires_a_positive_factor(self):
        with self.assertRaisesRegex(supabase_store.SupabaseStoreError, "maior que zero"):
            supabase_store._equivalence_factor("0")

    def test_equivalence_options_only_returns_active_groups(self):
        with patch.object(
            supabase_store,
            "_request",
            return_value=[{"grupo_id": "group-1"}, {"grupo_id": "group-2"}],
        ), patch.object(
            supabase_store,
            "get_equivalence_group",
            side_effect=[
                {"id": "group-1", "codigo": "EQ-ATIVO", "ativo": True, "members": []},
                {"id": "group-2", "codigo": "EQ-INATIVO", "ativo": False, "members": []},
            ],
        ):
            result = supabase_store.equivalence_options_for_sku("10220077")

        self.assertEqual(["EQ-ATIVO"], [group["codigo"] for group in result])


if __name__ == "__main__":
    unittest.main()
