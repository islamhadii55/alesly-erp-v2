import os
import re
import tempfile
import unittest
from pathlib import Path

DB_FILE = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
DB_FILE.close()
os.environ["DATABASE_PATH"] = DB_FILE.name
os.environ["APP_SECRET_KEY"] = "test-secret-key-that-is-long-enough-123456"
os.environ["INITIAL_ADMIN_PASSWORD"] = "test-admin-password-123456"
os.environ.pop("RAILWAY_ENVIRONMENT", None)
os.environ.pop("RAILWAY_PUBLIC_DOMAIN", None)
os.environ.pop("RAILWAY_PROJECT_ID", None)

from app import app, query  # noqa: E402


class AppSecurityTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        app.config.update(TESTING=True)
        cls.client = app.test_client()

    def csrf_token(self):
        response = self.client.get("/login")
        self.assertEqual(response.status_code, 200)
        match = re.search(r'name="csrf-token" content="([^"]+)"', response.get_data(as_text=True))
        self.assertIsNotNone(match)
        return match.group(1)

    def test_health_checks_database(self):
        response = self.client.get("/health")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json["database"], "ok")

    def test_post_without_csrf_is_rejected(self):
        response = self.client.post("/login", data={"username": "admin", "password": "test-admin-password-123456"})
        self.assertEqual(response.status_code, 400)

    def test_login_uses_hashed_password_and_csrf(self):
        token = self.csrf_token()
        response = self.client.post(
            "/login",
            data={"username": "admin", "password": "test-admin-password-123456", "csrf_token": token},
        )
        self.assertEqual(response.status_code, 302)
        with app.app_context():
            stored = query("SELECT password FROM users WHERE username='admin'", one=True)["password"]
        self.assertTrue(stored.startswith(("scrypt:", "pbkdf2:")))


if __name__ == "__main__":
    unittest.main()
