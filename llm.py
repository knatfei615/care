"""LLM integration – calls OpenAI to structure pharmacy monitoring notes."""

from __future__ import annotations

from openai import OpenAI

SYSTEM_PROMPT = """\
你是一位资深ICU临床药师，负责将ICU查房口述记录整理为标准药学监护记录。
你会收到“患者基本信息”和“本次查房口述记录”。患者基本信息包括年龄、性别、体重、入院日期、入院诊断等，是生成记录时必须结合的临床背景；不要要求用户在查房记录中重复提供这些内容。

输出格式（一行，不换行）：
主观资料：<与本次药学监护相关的症状、体征、病史、诊断等>。客观资料：<与本次药学监护相关的检验、检查结果等>。分析评估：<结合患者病理生理状态、疾病特点、用药情况及循证证据等进行分析评估>。药学监护建议：<个体化药物治疗方案建议、疗效和不良反应监护计划、药品不良反应识别与处理建议、患者用药指导等>。

规则：
1. 四段各以中文句号结尾，段落间无换行
2. 优先结合患者基本信息中的年龄、性别、体重、入院日期、入院诊断等内容判断疾病背景和用药风险
3. 保留原始输入中的药物名称、剂量、检验数值，不虚构未提及的数据
4. 如某段信息不足，用简短通用表述补充
5. 语言专业简洁
"""

_REQUIRED_MARKERS = ["主观资料：", "客观资料：", "分析评估：", "药学监护建议："]


def _validate(text: str) -> bool:
    """Return True if all four section markers are present."""
    return all(m in text for m in _REQUIRED_MARKERS)


def _build_user_content(patient_info: str, raw_text: str) -> str:
    """Build the user message with patient context before the free-text note."""
    patient_info = (patient_info or "").strip()
    raw_text = (raw_text or "").strip()

    if not patient_info:
        patient_info = "未提供。"

    return (
        "患者基本信息（生成记录时必须作为临床背景使用）：\n"
        f"{patient_info}\n\n"
        "本次查房口述记录：\n"
        f"{raw_text}"
    )


def structure_note(
    api_key: str,
    model: str,
    patient_info: str,
    raw_text: str,
    base_url: str = "",
) -> dict:
    """Call OpenAI-compatible API to structure *raw_text* into a standard note.

    Returns ``{"note": str, "error": str | None}``.
    """
    client = OpenAI(api_key=api_key, base_url=base_url or None)

    user_content = _build_user_content(patient_info, raw_text)

    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": user_content},
    ]

    for attempt in range(2):
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=messages,
                temperature=0.3,
                max_tokens=1000,
            )
            note = resp.choices[0].message.content.strip()
        except Exception as exc:
            return {"note": "", "error": f"OpenAI 调用失败：{exc}"}

        if _validate(note):
            return {"note": note, "error": None}

        # First attempt failed validation – retry with correction hint 2222
        if attempt == 0:
            messages.append({"role": "assistant", "content": note})
            messages.append({
                "role": "user",
                "content": (
                    "格式不正确，请严格按照要求重新输出。"
                    '必须包含"主观资料："、"客观资料："、"分析评估："、"药学监护建议："四个标记，'
                    "一行输出，不换行。"
                ),
            })

    # Both attempts failed validation – return what we have
    return {"note": note, "error": "格式校验未通过，请检查并手动修改。"}
