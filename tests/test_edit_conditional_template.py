import unittest

import excel_bancos
import main


class EditConditionalTemplateTests(unittest.TestCase):
    @staticmethod
    def _air_conditioning_fields(selected_supplier="1- GE"):
        return [
            {
                "key": "fornecedor",
                "label": "FORNECEDOR",
                "scope": "primaria",
                "selection_mode": "unitaria",
                "description_order": 1,
                "options": ["1- GE", "2- CLIM"],
                "selected_values": [selected_supplier],
                "selected_value": selected_supplier,
                "conjunto_only_options": [],
                "required": False,
                "free_text": False,
                "banco_mode": "",
            },
            {
                "key": "capacidade",
                "label": "CAPACIDADE",
                "scope": "secundaria",
                "selection_mode": "unitaria",
                "description_order": 2,
                "options": [],
                "selected_values": ["12000 BTU"],
                "selected_value": "12000 BTU",
                "conjunto_only_options": [],
                "required": False,
                "free_text": True,
                "banco_mode": "",
            },
        ]

    @staticmethod
    def _edit_context(fields):
        category = {"key": "cat_10_ar_condicionado", "label": "10 - AR CONDICIONADO"}
        return {
            "record": {
                "id": 1,
                "sku": "10100001",
                "descricao_primaria": "TESTE",
                "descricao_secundaria": "TESTE",
                "replacement_sku": "",
                "unidade": "pc",
                "possui_bom": False,
                "ativo": True,
            },
            "categories": [category],
            "pn_groups": [],
            "selected_category": category,
            "source_category": category,
            "selected_group_code": "10",
            "ordered_fields": fields,
            "conditional_rules": [],
            "workbook_path": "teste",
            "unit_options": ["pc"],
            "erro": "",
            "sucesso": "",
            "component_rows": [],
        }

    def test_air_conditioning_edit_uses_selects_and_keeps_free_text_fields(self):
        html = main.templates.env.get_template("editar_cadastro.html").render(
            **self._edit_context(self._air_conditioning_fields())
        )

        self.assertIn('<select name="fornecedor"', html)
        self.assertNotIn('<input name="fornecedor"', html)
        self.assertIn('<input name="capacidade"', html)

    def test_air_conditioning_edit_preserves_legacy_value_outside_current_options(self):
        html = main.templates.env.get_template("editar_cadastro.html").render(
            **self._edit_context(self._air_conditioning_fields("9- FORNECEDOR LEGADO"))
        )

        self.assertIn(
            '<option value="9- FORNECEDOR LEGADO" selected>9- FORNECEDOR LEGADO (valor atual)</option>',
            html,
        )

    def test_air_conditioning_create_uses_same_select_behavior_as_edit(self):
        category = {"key": "cat_10_ar_condicionado", "label": "10 - AR CONDICIONADO"}
        fields = self._air_conditioning_fields()
        html = main.templates.env.get_template("cadastro_bancos.html").render(
            categories=[category],
            selected_category=category,
            pn_groups=[],
            selected_group_code="10",
            selected_unit="pc",
            selected_bom_option="0",
            fields=fields,
            ordered_fields=fields,
            conditional_rules=[],
            workbook_path="teste",
            unit_options=["pc"],
            erro="",
            sucesso="",
            active_draft=None,
            supabase_mode=True,
            active_page="cadastro",
            component_rows=[],
        )

        self.assertIn('<select name="fornecedor"', html)
        self.assertNotIn('<input name="fornecedor"', html)
        self.assertIn('<input name="capacidade"', html)

    def test_any_category_uses_field_configuration_instead_of_category_exception(self):
        category = {"key": "cat_99_teste", "label": "99 - TESTE"}
        fields = self._air_conditioning_fields()
        context = self._edit_context(fields)
        context.update(
            selected_category=category,
            source_category=category,
            categories=[category],
        )

        html = main.templates.env.get_template("editar_cadastro.html").render(**context)

        self.assertIn('<select name="fornecedor"', html)
        self.assertNotIn('<input name="fornecedor"', html)
        self.assertIn('<input name="capacidade"', html)

        create_html = main.templates.env.get_template("cadastro_bancos.html").render(
            categories=[category],
            selected_category=category,
            pn_groups=[],
            selected_group_code="99",
            selected_unit="pc",
            selected_bom_option="0",
            fields=fields,
            ordered_fields=fields,
            conditional_rules=[],
            workbook_path="teste",
            unit_options=["pc"],
            erro="",
            sucesso="",
            active_draft=None,
            supabase_mode=True,
            active_page="cadastro",
            component_rows=[],
        )

        self.assertIn('<select name="fornecedor"', create_html)
        self.assertNotIn('<input name="fornecedor"', create_html)
        self.assertIn('<input name="capacidade"', create_html)

    def test_edit_template_loads_and_applies_bank_conditional_rules(self):
        values = {
            "braco": ["1- S/ AP BRACO"],
            "linha": ["1- LB"],
            "encosto": ["1- FIXO"],
            "pre_fixo": ["1- BCO"],
        }
        fields = excel_bancos.get_banco_fields_for_display("bancos")
        html = main.templates.env.get_template("editar_cadastro.html").render(
            record={
                "id": 1,
                "sku": "10200001",
                "descricao_primaria": "TESTE",
                "descricao_secundaria": "TESTE",
                "replacement_sku": "",
            },
            categories=[{"key": "bancos", "label": "Bancos"}],
            pn_groups=excel_bancos.list_pn_groups(),
            selected_category={"key": "bancos", "label": "Bancos"},
            source_category={"key": "bancos", "label": "Bancos"},
            selected_group_code="10",
            ordered_fields=main._enrich_fields(fields, values),
            conditional_rules=excel_bancos.get_conditional_rules_for_form("bancos"),
            workbook_path="teste",
            unit_options=["pc"],
            erro="",
            sucesso="",
        )

        self.assertIn('id="edit-primary-field-grid"', html)
        self.assertIn('id="edit-secondary-field-grid"', html)
        self.assertIn("const configuredRules =", html)
        self.assertIn('"sourceKey": "braco"', html)
        self.assertIn('"key": "lado_braco"', html)
        self.assertIn('"sourceLabel": "APOIO DE BRA\\u00c7O"', html)
        self.assertIn("normalizeOptionValue(value)", html)
        self.assertIn("rule.sourceKey", html)
        self.assertIn("target.key", html)
        self.assertIn("const updateConditionalFields = () =>", html)
        self.assertIn('if (action === "hide_option")', html)
        self.assertIn("target.optionValues", html)
        self.assertIn("option.hidden = isUnavailable", html)
        self.assertIn("updateConditionalFields();", html)

    def test_form_rules_keep_stable_keys_and_current_option_labels(self):
        rules = excel_bancos.get_conditional_rules_for_form("bancos")
        side_rules = [
            rule
            for rule in rules
            if rule.get("sourceKey") == "braco"
            and any(target.get("key") == "lado_braco" for target in rule.get("targets", []))
        ]

        self.assertTrue(side_rules)
        self.assertTrue(
            any(
                "1- S/ AP BRACO" in rule.get("values", [])
                or "SAPBRACO" in rule.get("values", [])
                for rule in side_rules
            )
        )


if __name__ == "__main__":
    unittest.main()
