import unittest

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


if __name__ == "__main__":
    unittest.main()
