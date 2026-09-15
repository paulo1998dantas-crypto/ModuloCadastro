import os
import unittest
from unittest.mock import patch

import main


class _FakeRequest:
    def __init__(self, values=None):
        self.values = values or {}

    async def form(self):
        return self

    def get(self, key, default=None):
        value = self.values.get(key, default)
        if isinstance(value, list):
            return value[-1] if value else default
        return value

    def getlist(self, key):
        value = self.values.get(key, [])
        return value if isinstance(value, list) else [value]


class BomCreationRouteTests(unittest.IsolatedAsyncioTestCase):
    def setUp(self):
        self.registration = {
            "id": 174,
            "sku": "10200176",
            "category_key": "bancos",
            "category_label": "20 - BANCOS",
            "descricao_primaria": "Banco de teste",
            "ativo": True,
        }

    def _common_patches(self):
        return (
            patch.dict(os.environ, {"CADASTRO_REQUIRE_LOGIN": "0"}),
            patch.object(main, "_supabase_mode", return_value=True),
            patch.object(main, "_workbook_display_path", return_value=""),
            patch.object(main.excel_bancos, "list_categories", return_value=[self.registration]),
            patch.object(
                main.excel_bancos,
                "selected_category",
                return_value={"key": "bancos", "label": "20 - BANCOS"},
            ),
        )

    async def test_get_opens_direct_creation_for_registered_item_without_bom(self):
        patches = self._common_patches()
        with patches[0], patches[1], patches[2], patches[3], patches[4], patch.object(
            main.supabase_store,
            "get_registration_by_sku",
            return_value=self.registration,
        ), patch.object(main.supabase_store, "get_bom_by_parent", return_value=None):
            response = await main.bom_novo_page(_FakeRequest(), item_pai="10200176")

        self.assertEqual(response.status_code, 200)
        body = response.body.decode("utf-8")
        self.assertIn("Criar nova B.O.M.", body)
        self.assertIn("10200176", body)
        self.assertIn('action="/bom/novo"', body)

    async def test_get_redirects_to_existing_bom_instead_of_creating_duplicate(self):
        patches = self._common_patches()
        with patches[0], patches[1], patches[2], patches[3], patches[4], patch.object(
            main.supabase_store,
            "get_registration_by_sku",
            return_value=self.registration,
        ), patch.object(
            main.supabase_store,
            "get_bom_by_parent",
            return_value={"header": {"id": 473}},
        ):
            response = await main.bom_novo_page(_FakeRequest(), item_pai="10200176")

        self.assertEqual(response.status_code, 303)
        self.assertTrue(response.headers["location"].startswith("/bom/473?sucesso="))

    async def test_post_creates_bom_from_registered_item_and_components(self):
        patches = self._common_patches()
        with patches[0], patches[1], patches[2], patches[3], patches[4], patch.object(
            main.supabase_store,
            "get_registration_by_sku",
            return_value=self.registration,
        ), patch.object(main.supabase_store, "get_bom_by_parent", return_value=None), patch.object(
            main.supabase_store,
            "save_bom",
            return_value={"bom": {"id": 999}, "components_count": 1},
        ) as save_bom:
            response = await main.bom_novo_post(
                _FakeRequest(
                    {
                        "parent_sku": "10200176",
                        "parent_descricao": "Banco de teste",
                        "component_search": "10220023 - Componente",
                        "component_codigo": "10220023",
                        "component_descricao": "Componente",
                        "component_unidade": "pc",
                        "component_quantidade": "1",
                    }
                )
            )

        self.assertEqual(response.status_code, 303)
        self.assertTrue(response.headers["location"].startswith("/bom/999?sucesso="))
        save_bom.assert_called_once()
        self.assertEqual(save_bom.call_args.args[0], "10200176")
        self.assertEqual(save_bom.call_args.args[2][0]["codigo"], "10220023")

    async def test_get_shows_error_for_unknown_parent_sku(self):
        patches = self._common_patches()
        with patches[0], patches[1], patches[2], patches[3], patches[4], patch.object(
            main.supabase_store,
            "get_registration_by_sku",
            return_value=None,
        ):
            response = await main.bom_novo_page(_FakeRequest(), item_pai="99999999")

        self.assertEqual(response.status_code, 200)
        self.assertIn("Item pai não encontrado", response.body.decode("utf-8"))


if __name__ == "__main__":
    unittest.main()
