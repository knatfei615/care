import unittest
import re
from pathlib import Path


class FrontendApiParsingTest(unittest.TestCase):
    def test_generate_note_uses_non_json_safe_parser(self):
        template = Path("templates/index.html").read_text(encoding="utf-8")
        generate_note_match = re.search(
            r"async function generateNote\(\) \{.*?/\* ── Save ── \*/",
            template,
            re.S,
        )

        self.assertIn("async function readApiJson", template)
        self.assertIsNotNone(generate_note_match)
        self.assertIn("const data = await readApiJson(res);", generate_note_match.group(0))


if __name__ == "__main__":
    unittest.main()
