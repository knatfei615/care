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

    def test_selecting_new_patient_defaults_note_date_to_admission_date(self):
        template = Path("templates/index.html").read_text(encoding="utf-8")
        select_patient_match = re.search(
            r"function selectPatient\(rowIdx\) \{.*?async function loadSlotStatuses",
            template,
            re.S,
        )
        helper_match = re.search(
            r"function applyDefaultRecordDate\(patient\) \{.*?\n\}",
            template,
            re.S,
        )

        self.assertIsNotNone(select_patient_match)
        self.assertIsNotNone(helper_match)
        self.assertIn("applyDefaultRecordDate(selectedPatient);", select_patient_match.group(0))
        self.assertIn('patient.status === "no_record"', helper_match.group(0))
        self.assertIn("patient.admission_date", helper_match.group(0))


if __name__ == "__main__":
    unittest.main()
