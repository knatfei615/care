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
        self.assertIn("本次要求输出的药学监护记录：", content)
        self.assertLess(content.index("患者基本信息"), content.index("本次要求输出的药学监护记录"))

    def test_includes_prior_notes_between_patient_context_and_current_note(self):
        content = _build_user_content(
            patient_info="年龄：3岁，诊断：重症肺炎",
            raw_text="今日体温下降，继续评估抗感染疗效。",
            prior_notes="记录1：日期：2026-04-28；分级：一级监护；类型：药学监护；内容：昨日建议关注万古霉素谷浓度。",
        )

        self.assertIn("既往药学监护记录（仅作连续性参考，不要原文照抄）：", content)
        self.assertIn("昨日建议关注万古霉素谷浓度", content)
        self.assertLess(content.index("患者基本信息"), content.index("既往药学监护记录"))
        self.assertLess(content.index("既往药学监护记录"), content.index("本次要求输出的药学监护记录"))

    def test_includes_current_medications_between_prior_notes_and_current_note(self):
        content = _build_user_content(
            patient_info="年龄：3岁，诊断：重症肺炎",
            raw_text="今日复评抗感染方案。",
            prior_notes="记录1：昨日建议关注肾功能。",
            current_medications="万古霉素 1g q12h IV\n美罗培南 0.5g q8h IV",
        )

        self.assertIn("当前药物医嘱（用户手动维护的最新医嘱，作为本次评估的用药背景）：", content)
        self.assertIn("万古霉素 1g q12h IV", content)
        self.assertLess(content.index("既往药学监护记录"), content.index("当前药物医嘱"))
        self.assertLess(content.index("当前药物医嘱"), content.index("本次要求输出的药学监护记录"))

    def test_omits_prior_notes_section_when_empty(self):
        content = _build_user_content(
            patient_info="年龄：3岁，诊断：重症肺炎",
            raw_text="今日复评。",
            prior_notes="",
        )

        self.assertNotIn("既往药学监护记录", content)


if __name__ == "__main__":
    unittest.main()
