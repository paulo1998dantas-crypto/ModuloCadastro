import unittest
from copy import deepcopy
from unittest.mock import patch

import excel_bancos


def _catalog():
    return {
        "active_category": "teste",
        "pn_groups": [
            {"code": "20", "label": "PRODUTO / PROCESSO", "prefixes": ["PP"]},
            {"code": "30", "label": "CONJUNTO / KIT", "prefixes": ["CJ"]},
        ],
        "categories": [
            {
                "key": "teste",
                "label": "10 - TESTE",
                "fields": [
                    {
                        "key": "origem",
                        "label": "ORIGEM",
                        "scope": "primaria",
                        "selection_mode": "unitaria",
                        "options": ["1- ATIVO", "2- INATIVO"],
                    },
                    {
                        "key": "destino_a",
                        "label": "DESTINO A",
                        "scope": "primaria",
                        "selection_mode": "unitaria",
                        "options": [],
                    },
                    {
                        "key": "destino_b",
                        "label": "DESTINO B",
                        "scope": "secundaria",
                        "selection_mode": "unitaria",
                        "options": [],
                    },
                ],
                "conditional_rules": [],
            }
        ],
    }


class ConditionalRuleBulkTests(unittest.TestCase):
    def test_multiple_trigger_values_and_targets_are_saved_atomically(self):
        catalog = _catalog()
        with patch.object(excel_bancos, "load_catalog", return_value=catalog), patch.object(
            excel_bancos, "save_catalog"
        ) as save_catalog:
            result = excel_bancos.add_conditional_rules(
                "teste",
                "origem",
                ["1- ATIVO", "2- INATIVO"],
                ["destino_a", "destino_b"],
                action_value="hide",
            )

        rules = catalog["categories"][0]["conditional_rules"]
        self.assertEqual(len(rules), 2)
        self.assertEqual(len(result["rules"]), 2)
        self.assertEqual({rule["target_field_key"] for rule in rules}, {"destino_a", "destino_b"})
        self.assertTrue(all(rule["source_values"] == ["ATIVO", "INATIVO"] for rule in rules))
        save_catalog.assert_called_once_with(catalog)

    def test_duplicate_in_bulk_selection_does_not_partially_save(self):
        catalog = _catalog()
        catalog["categories"][0]["conditional_rules"].append(
            {
                "key": "existente",
                "source_type": "field",
                "source_field_key": "origem",
                "source_field_label": "ORIGEM",
                "source_field_scope": "primaria",
                "source_values": ["ATIVO"],
                "target_field_key": "destino_b",
                "target_field_label": "DESTINO B",
                "target_field_scope": "secundaria",
                "action": "hide",
            }
        )
        original = deepcopy(catalog["categories"][0]["conditional_rules"])

        with patch.object(excel_bancos, "load_catalog", return_value=catalog), patch.object(
            excel_bancos, "save_catalog"
        ) as save_catalog:
            with self.assertRaisesRegex(ValueError, "j.*existe"):
                excel_bancos.add_conditional_rules(
                    "teste",
                    "origem",
                    ["ATIVO", "INATIVO"],
                    ["destino_a", "destino_b"],
                    action_value="hide",
                )

        self.assertEqual(catalog["categories"][0]["conditional_rules"], original)
        save_catalog.assert_not_called()


if __name__ == "__main__":
    unittest.main()
