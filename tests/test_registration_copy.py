import unittest
from unittest.mock import patch

import main
import supabase_store


class RegistrationCopyTests(unittest.TestCase):
    def setUp(self):
        self.source = {
            "id": 42,
            "sku": "10200176",
            "category_key": "cat_10_ar_condicionado",
            "category_label": "10 - AR CONDICIONADO",
            "descricao_primaria": "EVAPORADOR DUPLO",
            "unidade": "CJ",
            "possui_bom": True,
        }
        self.editable = {
            "record": self.source,
            "category": {"key": "cat_10_ar_condicionado", "label": "10 - AR CONDICIONADO"},
            "source_category": {"key": "cat_10_ar_condicionado", "label": "10 - AR CONDICIONADO"},
            "current_group_code": "10",
            "fields": [],
            "groups": {"grupo_codigo": ["10"], "descritor_base": ["1- EVAPORADOR DUPLO"]},
        }

    def test_copy_prepares_new_form_groups_without_reusing_sku(self):
        bom = {
            "header": {"id": 9, "parent_sku": "10200176"},
            "components": [
                {
                    "component_sku": "10220023",
                    "component_descricao": "COMPONENTE",
                    "unidade": "PC",
                    "quantidade": 2,
                }
            ],
        }
        with patch.object(supabase_store, "editable_registration", return_value=self.editable), patch.object(
            supabase_store, "get_bom_by_parent", return_value=bom
        ):
            result = supabase_store.copy_registration_for_new(42)

        self.assertNotIn("sku", result["groups"])
        self.assertEqual(result["groups"]["unidade"], ["cj"])
        self.assertEqual(result["groups"]["possui_bom"], ["1"])
        self.assertEqual(result["groups"]["component_codigo"], ["10220023"])
        self.assertEqual(result["groups"]["component_quantidade"], ["2"])

    def test_copy_without_bom_stays_valid_for_new_cadastro(self):
        editable = {**self.editable, "record": {**self.source, "possui_bom": False}}
        with patch.object(supabase_store, "editable_registration", return_value=editable), patch.object(
            supabase_store, "get_bom_by_parent", return_value=None
        ):
            result = supabase_store.copy_registration_for_new(42)

        self.assertEqual(result["groups"]["possui_bom"], ["0"])
        self.assertNotIn("component_codigo", result["groups"])

    def test_new_cadastro_route_loads_copy_before_rendering(self):
        rendered = object()
        copied = {
            **self.editable,
            "groups": {"grupo_codigo": ["10"]},
        }
        request = object()
        with patch.object(main, "_supabase_mode", return_value=True), patch.object(
            main.supabase_store, "copy_registration_for_new", return_value=copied
        ) as copy_registration, patch.object(main, "_render_cadastro_page", return_value=rendered) as render:
            result = __import__("asyncio").run(
                main.cadastro_bancos_page(request, categoria="", copiar_de="42")
            )

        self.assertIs(result, rendered)
        copy_registration.assert_called_once_with("42")
        render.assert_called_once_with(
            request,
            categoria="cat_10_ar_condicionado",
            sucesso="",
            erro="",
            draft_id="",
            copy_source=self.source,
            initial_groups={"grupo_codigo": ["10"]},
        )

    def test_search_registration_bases_returns_only_form_selection_data(self):
        with patch.object(
            supabase_store,
            "list_registrations",
            return_value=[
                {
                    "id": 42,
                    "sku": "10200176",
                    "descricao_primaria": "EVAPORADOR DUPLO",
                    "category_label": "10 - AR CONDICIONADO",
                    "category_key": "cat_10_ar_condicionado",
                    "unidade": "CJ",
                }
            ],
        ) as list_registrations:
            result = supabase_store.search_registration_bases("evaporador", "cat_10_ar_condicionado")

        list_registrations.assert_called_once_with(
            category_key="cat_10_ar_condicionado",
            query="evaporador",
            include_inactive=False,
            limit=25,
        )
        self.assertEqual(result[0]["id"], 42)
        self.assertEqual(result[0]["sku"], "10200176")
        self.assertEqual(result[0]["unidade"], "cj")


if __name__ == "__main__":
    unittest.main()
