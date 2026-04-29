import unittest

from llm import _build_user_content


class BuildUserContentTest(unittest.TestCase):
    def test_includes_patient_context_before_rounding_note(self):
        content = _build_user_content(
            patient_info="年龄：3岁，性别：男，体重：14kg，诊断：重症肺炎",
            raw_text="今天复查感染指标下降，继续当前抗感染方案。",
        )

        self.assertIn("患者基本信息（生成记录时必须作为临床背景使用）：", content)
        self.assertIn("年龄：3岁", content)
        self.assertIn("诊断：重症肺炎", content)
        self.assertIn("本次查房口述记录：", content)
        self.assertLess(content.index("患者基本信息"), content.index("本次查房口述记录"))


if __name__ == "__main__":
    unittest.main()
