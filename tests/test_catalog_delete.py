from types import SimpleNamespace
import unittest
from unittest.mock import patch

import main
import supabase_store


class CatalogDeleteAuthorizationTests(unittest.TestCase):
    def test_only_admin_role_can_delete_with_shared_rbac(self):
        request = SimpleNamespace(
            state=SimpleNamespace(
                erp_access={
                    "active": True,
                    "roles": ["ADMIN", "ENGENHARIA"],
                    "permissions": ["cadastro.access"],
                }
            )
        )
        self.assertTrue(main._cadastro_delete_allowed(request))

    def test_engineering_cadastro_permission_cannot_delete(self):
        request = SimpleNamespace(
            state=SimpleNamespace(
                erp_access={
                    "active": True,
                    "roles": ["ENGENHARIA"],
                    "permissions": ["cadastro.access"],
                }
            )
        )
        self.assertFalse(main._cadastro_delete_allowed(request))

    def test_inactive_admin_cannot_delete(self):
        request = SimpleNamespace(
            state=SimpleNamespace(
                erp_access={"active": False, "roles": ["ADMIN"], "permissions": []}
            )
        )
        self.assertFalse(main._cadastro_delete_allowed(request))


class CatalogDeleteStoreTests(unittest.TestCase):
    def test_delete_calls_only_protected_rpc(self):
        response = {"deleted": False, "sku": "30200049", "blockers": ["B.O.M. operacional"]}
        with patch.object(supabase_store, "_request", return_value=response) as request:
            self.assertEqual(response, supabase_store.delete_registration(81))
        request.assert_called_once_with(
            "POST",
            "rpc/erp_delete_catalog_sku",
            payload={"p_registration_id": 81},
        )

    def test_delete_rejects_invalid_id_without_database_call(self):
        with patch.object(supabase_store, "_request") as request:
            with self.assertRaises(supabase_store.SupabaseStoreError):
                supabase_store.delete_registration("x")
        request.assert_not_called()


if __name__ == "__main__":
    unittest.main()
