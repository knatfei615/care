import json
import unittest
from pathlib import Path


class DeploymentConfigTest(unittest.TestCase):
    def test_zeabur_startup_sets_gunicorn_timeout(self):
        config = json.loads(Path("zbpack.json").read_text(encoding="utf-8"))

        start_command = config["start_command"]

        self.assertIn("_startup", start_command)
        self.assertIn("GUNICORN_CMD_ARGS", start_command)
        self.assertIn("--timeout", start_command)
        self.assertIn("120", start_command)


if __name__ == "__main__":
    unittest.main()
