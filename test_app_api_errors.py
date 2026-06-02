import unittest

from app import app


class ApiErrorResponseTest(unittest.TestCase):
    def setUp(self):
        self.original_testing = app.config.get("TESTING")
        self.original_csrf_enabled = app.config.get("WTF_CSRF_ENABLED", True)
        app.config["TESTING"] = True

    def tearDown(self):
        app.config["TESTING"] = self.original_testing
        app.config["WTF_CSRF_ENABLED"] = self.original_csrf_enabled

    def test_api_csrf_failure_returns_json(self):
        app.config["WTF_CSRF_ENABLED"] = True
        client = app.test_client()

        response = client.post("/api/generate", json={"raw_text": "测试"})

        self.assertEqual(response.status_code, 400)
        self.assertEqual(response.mimetype, "application/json")
        payload = response.get_json()
        self.assertEqual(payload["error_type"], "csrf_error")
        self.assertIn("刷新页面", payload["recovery_hint"])

    def test_api_login_required_returns_json_instead_of_redirect(self):
        app.config["WTF_CSRF_ENABLED"] = False
        client = app.test_client()

        response = client.post("/api/generate", json={"raw_text": "测试"})

        self.assertEqual(response.status_code, 401)
        self.assertEqual(response.mimetype, "application/json")
        self.assertIsNone(response.headers.get("Location"))
        payload = response.get_json()
        self.assertEqual(payload["error_type"], "auth_error")
        self.assertIn("重新登录", payload["recovery_hint"])


if __name__ == "__main__":
    unittest.main()
