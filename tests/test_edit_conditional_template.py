import unittest

import excel_bancos
import main


class EditConditionalTemplateTests(unittest.TestCase):
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
        self.assertIn('"sourceLabel": "APOIO DE BRA\\u00c7O"', html)
        self.assertIn("const updateConditionalFields = () =>", html)
        self.assertIn("updateConditionalFields();", html)


if __name__ == "__main__":
    unittest.main()
