from types import SimpleNamespace
import unittest
from unittest.mock import patch

import main


class SharedRbacTest(unittest.TestCase):
    def test_cadastro_access_requires_admin_or_permission(self):
        self.assertTrue(
            main._cadastro_access_allowed(
                {
                    "active": True,
                    "roles": ["ENGENHARIA"],
                    "permissions": ["cadastro.access"],
                }
            )
        )
        self.assertTrue(
            main._cadastro_access_allowed(
                {"active": True, "roles": ["ADMIN"], "permissions": []}
            )
        )
        self.assertFalse(
            main._cadastro_access_allowed(
                {"active": True, "roles": ["PCP"], "permissions": []}
            )
        )
        self.assertFalse(
            main._cadastro_access_allowed(
                {
                    "active": False,
                    "roles": ["ADMIN"],
                    "permissions": ["cadastro.access"],
                }
            )
        )

    def test_shared_login_rejects_user_without_cadastro_permission(self):
        user = {
            "id": 7,
            "username": "pcp",
            "password_hash": "hash",
            "active": True,
            "auth_version": 1,
        }
        with (
            patch.object(main, "_shared_user_lookup", return_value=user),
            patch.object(main, "check_password_hash", return_value=True),
            patch.object(main, "_shared_rbac_enabled", return_value=True),
            patch.object(
                main,
                "_shared_access_lookup",
                return_value={
                    "active": True,
                    "auth_version": 1,
                    "roles": ["PCP"],
                    "permissions": [],
                },
            ),
        ):
            self.assertIsNone(main._authenticate_shared_user("pcp", "senha"))

    def test_legacy_bridge_never_grants_cadastro_to_operator(self):
        user = {
            "id": 7,
            "username": "operador",
            "password_hash": "hash",
            "role": "OPERADOR",
            "active": True,
        }
        with (
            patch.object(main, "_shared_user_lookup", return_value=user),
            patch.object(main, "check_password_hash", return_value=True),
            patch.object(main, "_shared_rbac_enabled", return_value=False),
        ):
            self.assertIsNone(
                main._authenticate_shared_user("operador", "senha")
            )

    def test_legacy_bridge_keeps_admin_and_engineering_access(self):
        for role in ("ADM", "ADMIN", "ENGENHARIA", "PCP"):
            user = {
                "id": 7,
                "username": role.lower(),
                "password_hash": "hash",
                "role": role,
                "active": True,
            }
            with (
                self.subTest(role=role),
                patch.object(main, "_shared_user_lookup", return_value=user),
                patch.object(main, "check_password_hash", return_value=True),
                patch.object(main, "_shared_rbac_enabled", return_value=False),
            ):
                self.assertIsNotNone(
                    main._authenticate_shared_user(role.lower(), "senha")
                )

    def test_pcp_has_read_access_but_not_write_access(self):
        request = SimpleNamespace(
            state=SimpleNamespace(
                erp_access={"active": True, "roles": ["PCP"], "permissions": ["cadastro.access"]}
            )
        )
        self.assertTrue(main._cadastro_access_allowed(request.state.erp_access))
        self.assertFalse(main._cadastro_write_allowed(request))

    def test_engineering_can_maintain_cadastro(self):
        request = SimpleNamespace(
            state=SimpleNamespace(
                erp_access={"active": True, "roles": ["ENGENHARIA"], "permissions": ["cadastro.access"]}
            )
        )
        self.assertTrue(main._cadastro_write_allowed(request))

    def test_session_payload_contains_auth_version_and_redirect_is_local(self):
        with patch.object(main, "_session_secret", return_value=b"test-secret"):
            raw = main._make_session("engenharia", user_id=9, auth_version=4)
            request = SimpleNamespace(cookies={main.SESSION_COOKIE: raw})
            payload = main._read_session_payload(request)
        self.assertEqual("engenharia", payload["u"])
        self.assertEqual(9, payload["uid"])
        self.assertEqual(4, payload["av"])
        self.assertEqual(
            "/cadastros?tab=sku",
            main._safe_next_path("/cadastros?tab=sku"),
        )
        self.assertEqual(
            "/cadastro/bancos",
            main._safe_next_path("https://evil.example/"),
        )
        self.assertEqual(
            "/cadastro/bancos",
            main._safe_next_path("//evil.example/"),
        )


if __name__ == "__main__":
    unittest.main()
