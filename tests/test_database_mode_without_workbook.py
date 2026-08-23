import os
import unittest
from unittest.mock import patch

import excel_bancos


class DatabaseModeWithoutWorkbookTests(unittest.TestCase):
    def test_supabase_mode_does_not_sync_legacy_workbook(self):
        with (
            patch.dict(os.environ, {"CADASTRO_SAVE_MODE": "supabase"}, clear=False),
            patch.object(excel_bancos, "active_workbook_path") as workbook_path,
            patch.object(excel_bancos, "sync_workbook_headers") as sync_headers,
        ):
            excel_bancos.sync_workbook_structure("cat_10_ar_condicionado")

        workbook_path.assert_not_called()
        sync_headers.assert_not_called()

    def test_local_mode_keeps_optional_workbook_compatibility(self):
        with (
            patch.dict(os.environ, {"CADASTRO_SAVE_MODE": "local"}, clear=False),
            patch.object(excel_bancos, "active_workbook_path", return_value="cadastro.xlsx") as workbook_path,
            patch.object(excel_bancos, "sync_workbook_headers") as sync_headers,
        ):
            excel_bancos.sync_workbook_structure("cat_10_ar_condicionado")

        workbook_path.assert_called_once_with()
        sync_headers.assert_called_once_with("cadastro.xlsx", "cat_10_ar_condicionado")


if __name__ == "__main__":
    unittest.main()
