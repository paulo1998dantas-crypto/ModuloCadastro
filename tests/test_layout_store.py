import unittest
from unittest.mock import patch

import supabase_store


class LayoutStoreTests(unittest.TestCase):
    def test_list_layouts_adds_os_usage_count(self):
        layouts = [{"id": "layout-a", "updated_at": "2026-08-13T10:00:00"}]
        documents = [
            {"layout_arquivo_id": "layout-a"},
            {"layout_arquivo_id": "layout-a"},
            {"layout_arquivo_id": "layout-b"},
        ]
        with patch.object(supabase_store, "_request_all", side_effect=[layouts, documents]):
            result = supabase_store.list_layouts()

        self.assertEqual(2, result[0]["uso_os"])

    def test_download_layout_only_uses_private_layout_bucket(self):
        layout = {
            "id": "layout-a",
            "storage_bucket": "os-layouts",
            "storage_path": "layouts/hash.pdf",
            "mime_type": "application/pdf",
        }
        with (
            patch.object(supabase_store, "get_layout", return_value=layout),
            patch.object(supabase_store, "_storage_request", return_value=(b"%PDF-1.4", {})),
        ):
            content, returned = supabase_store.download_layout("layout-a")

        self.assertEqual(b"%PDF-1.4", content)
        self.assertEqual("layout-a", returned["id"])


if __name__ == "__main__":
    unittest.main()
