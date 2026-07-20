import unittest
from unittest.mock import patch

import excel_bancos
import supabase_store


class BomPreferenceTests(unittest.TestCase):
    def test_explicit_yes_does_not_depend_on_prefix(self):
        self.assertTrue(excel_bancos.requires_component_bom([], {"possui_bom": "1"}))

    def test_explicit_no_overrides_cj_prefix(self):
        fields = [{"key": "prefixo", "label": "PREFIXO", "scope": "primaria"}]
        data = {"prefixo": "CJ", "possui_bom": "0"}

        self.assertFalse(excel_bancos.requires_component_bom(fields, data))

    def test_missing_choice_no_longer_infers_from_prefix(self):
        fields = [{"key": "prefixo", "label": "PREFIXO", "scope": "primaria"}]

        self.assertFalse(excel_bancos.requires_component_bom(fields, {"prefixo": "PP"}))

    def test_stored_preference_preserves_boolean_values(self):
        self.assertTrue(supabase_store._stored_bom_preference({"form_values": {"possui_bom": True}}))
        self.assertFalse(supabase_store._stored_bom_preference({"form_values": {"possui_bom": False}}))

    def test_missing_stored_preference_is_undefined(self):
        self.assertIsNone(supabase_store._stored_bom_preference({"form_values": {}}))

    def test_legacy_veiculo_pb_defaults_to_bom(self):
        row = {
            "category_key": "cat_34_veiculo_p_b",
            "form_values": {},
        }

        self.assertTrue(supabase_store._stored_bom_preference(row))

    def test_explicit_veiculo_pb_preference_is_preserved(self):
        row = {
            "category_key": "cat_34_veiculo_p_b",
            "form_values": {"possui_bom": False},
        }

        self.assertFalse(supabase_store._stored_bom_preference(row))


class SkuStructureMigrationTests(unittest.TestCase):
    def test_category_alias_does_not_trigger_structure_migration(self):
        current = {"category_key": "cat_20_bco", "sku": "30200012", "form_values": {"grupo_codigo": ["30"]}}
        category = {"key": "bancos", "label": "20 - BANCOS"}

        self.assertFalse(
            supabase_store._registration_structure_changed(
                current,
                category,
                [],
                {"grupo_codigo": "30"},
            )
        )

    def test_structure_change_compares_group_and_category(self):
        current = {"category_key": "cat_14_piso", "sku": "30140027"}
        category = {"key": "cat_14_piso", "label": "14 - PISO"}

        self.assertFalse(
            supabase_store._registration_structure_changed(
                current,
                category,
                [],
                {"grupo_codigo": "30"},
            )
        )
        self.assertTrue(
            supabase_store._registration_structure_changed(
                current,
                category,
                [],
                {"grupo_codigo": "20"},
            )
        )
        self.assertTrue(
            supabase_store._registration_structure_changed(
                current,
                {"key": "cat_16_isolamento", "label": "16 - ISOLAMENTO"},
                [],
                {"grupo_codigo": "30"},
            )
        )

    def test_bom_migration_updates_parent_and_child_references(self):
        snapshots = {
            "headers": [
                {
                    "id": 5,
                    "parent_sku": "30140027",
                    "source": "cadastro",
                }
            ],
            "components": [
                {
                    "id": 7,
                    "bom_id": 5,
                    "parent_sku": "30140027",
                    "component_sku": "10300001",
                    "component_descricao": "FIXADOR",
                    "unidade": "pc",
                    "quantidade": 1,
                    "ordem": 1,
                    "search_text": "",
                },
                {
                    "id": 8,
                    "bom_id": 9,
                    "parent_sku": "40340049",
                    "component_sku": "30140027",
                    "component_descricao": "PISO ANTIGO",
                    "unidade": "cj",
                    "quantidade": 1,
                    "ordem": 2,
                    "search_text": "",
                },
            ],
        }
        new_record = {
            "id": 99,
            "sku": "20140031",
            "category_key": "cat_14_piso",
            "category_label": "14 - PISO",
            "descricao_primaria": "PP PISO CORRIGIDO",
            "unidade": "cj",
        }

        with patch.object(supabase_store, "_request", return_value=None) as request:
            result = supabase_store._apply_bom_sku_migration(
                snapshots,
                "30140027",
                new_record,
            )

        self.assertEqual(result, {"bom_headers": 1, "bom_components": 2})
        payloads_by_id = {
            int(call.args[2][0][1].removeprefix("eq.")): call.kwargs["payload"]
            for call in request.call_args_list
        }
        self.assertEqual(payloads_by_id[5]["parent_sku"], "20140031")
        self.assertEqual(payloads_by_id[5]["registration_id"], 99)
        self.assertEqual(payloads_by_id[7]["parent_sku"], "20140031")
        self.assertEqual(payloads_by_id[8]["component_sku"], "20140031")
        self.assertEqual(payloads_by_id[8]["component_descricao"], "PP PISO CORRIGIDO")

        with patch.object(supabase_store, "_request", return_value=None) as restore_request:
            supabase_store._restore_bom_references(snapshots)
        restored_by_id = {
            int(call.args[2][0][1].removeprefix("eq.")): call.kwargs["payload"]
            for call in restore_request.call_args_list
        }
        self.assertEqual(restored_by_id[5]["parent_sku"], "30140027")
        self.assertEqual(restored_by_id[7]["parent_sku"], "30140027")
        self.assertEqual(restored_by_id[8]["component_sku"], "30140027")

    def test_update_creates_replacement_and_inactivates_old_sku(self):
        current = {
            "id": 12,
            "category_key": "cat_14_piso",
            "category_label": "14 - PISO",
            "sku": "30140027",
            "ativo": True,
            "form_values": {"possui_bom": True},
            "search_text": "30140027 PISO",
        }
        category = {"key": "cat_14_piso", "label": "14 - PISO"}
        payload = {
            "category_key": category["key"],
            "category_label": category["label"],
            "sku": "20140031",
            "descricao_primaria": "PP PISO CORRIGIDO",
            "descricao_secundaria": "PP PISO CORRIGIDO COMPLETO",
            "unidade": "cj",
            "ativo": True,
            "form_values": {"grupo_codigo": ["20"], "possui_bom": True},
        }
        new_record = {**payload, "id": 44}

        def request_side_effect(method, table, query=None, payload=None, prefer=""):
            if method == "POST" and table == supabase_store.REGISTRATIONS_TABLE:
                return [new_record]
            if method == "PATCH" and table == supabase_store.REGISTRATIONS_TABLE:
                return [{**current, **(payload or {})}]
            return None

        with (
            patch.object(supabase_store, "get_registration", return_value=current),
            patch.object(supabase_store, "_category", return_value=category),
            patch.object(excel_bancos, "get_banco_fields", return_value=[]),
            patch.object(supabase_store, "_next_sku", return_value="20140031"),
            patch.object(
                supabase_store,
                "_registration_payload",
                return_value=(payload, {"primaria": payload["descricao_primaria"], "secundaria": payload["descricao_secundaria"]}, True),
            ),
            patch.object(supabase_store, "_duplicate_exists", return_value=False),
            patch.object(supabase_store, "_bom_reference_snapshots", return_value={"headers": [], "components": []}),
            patch.object(
                supabase_store,
                "_apply_bom_sku_migration",
                return_value={"bom_headers": 0, "bom_components": 0},
            ),
            patch.object(supabase_store, "_request", side_effect=request_side_effect) as request,
        ):
            result = supabase_store.update_registration(
                12,
                {
                    "categoria": "cat_14_piso",
                    "grupo_codigo": "20",
                    "confirmar_migracao": "1",
                },
            )

        self.assertTrue(result["migrated"])
        self.assertEqual(result["previous_sku"], "30140027")
        self.assertEqual(result["sku"], "20140031")
        old_patch = next(
            call
            for call in request.call_args_list
            if call.args[0] == "PATCH" and call.args[1] == supabase_store.REGISTRATIONS_TABLE
        )
        self.assertFalse(old_patch.kwargs["payload"]["ativo"])
        migration = old_patch.kwargs["payload"]["form_values"][supabase_store.SKU_MIGRATION_FORM_KEY]
        self.assertEqual(migration["replacement_sku"], "20140031")


class CategoryAliasTests(unittest.TestCase):
    def test_sanitizer_merges_cj_bco_fields_into_bancos(self):
        catalog = {
            "active_category": "cat_20_bco",
            "categories": [
                {
                    "key": "bancos",
                    "label": "20 - BANCOS",
                    "sheet_name": "20 - BANCOS",
                    "fields": [{"key": "pre_fixo", "label": "PRE FIXO", "scope": "primaria", "options": ["1- BCO"]}],
                },
                {
                    "key": "cat_20_bco",
                    "label": "20 - CJ-BCO",
                    "sheet_name": "20 - CJ-BCO",
                    "fields": [{"key": "prefixo", "label": "PREFIXO", "scope": "primaria", "options": ["CJ"]}],
                },
            ],
        }

        sanitized = excel_bancos._sanitize_catalog(catalog)

        self.assertEqual(sanitized["active_category"], "bancos")
        self.assertEqual([category["label"] for category in sanitized["categories"]], ["20 - BANCOS"])
        self.assertEqual(
            [field["key"] for field in sanitized["categories"][0]["fields"]],
            ["pre_fixo", "prefixo"],
        )

    def test_supabase_category_filter_includes_legacy_aliases(self):
        self.assertEqual(
            supabase_store._category_key_filter("category_key", "bancos"),
            ("category_key", "in.(bancos,cat_20_bco,cat_20_cj_bco,20_cj_bco)"),
        )
        self.assertEqual(
            supabase_store._category_key_filter("category_key", "cat_20_bco"),
            ("category_key", "in.(bancos,cat_20_bco,cat_20_cj_bco,20_cj_bco)"),
        )


if __name__ == "__main__":
    unittest.main()
