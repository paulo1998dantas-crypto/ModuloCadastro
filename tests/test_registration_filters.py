import unittest
from unittest.mock import patch

import supabase_store


class RegistrationFilterTests(unittest.TestCase):
    def test_layout_filter_preserves_commas(self):
        self.assertEqual("3,2,3", supabase_store._safe_filter_value("3,2,3", "cj_layout"))

    def test_layout_filter_accepts_parentheses_and_semicolons(self):
        self.assertEqual("3,2,3", supabase_store._safe_filter_value("(3, 2, 3)", "cj_layout"))
        self.assertEqual("3,2,3", supabase_store._safe_filter_value("3;2;3", "cj_layout"))

    def test_list_registrations_sends_normalized_layout_to_supabase(self):
        with patch.object(supabase_store, "_request", return_value=[]) as request:
            supabase_store.list_registrations(
                "bancos",
                filters={"cj_layout": "(3,2,3)"},
            )

        params = request.call_args.args[2]
        self.assertIn(("field_values->>cj_layout", "ilike.*3,2,3*"), params)


if __name__ == "__main__":
    unittest.main()
