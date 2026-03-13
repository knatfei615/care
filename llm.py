"""LLM integration – calls OpenAI to structure pharmacy monitoring notes."""

from __future__ import annotations

from openai import OpenAI

SYSTEM_PROMPT = """\
你是一位资深 PICU 临床药师，负责将查房口述记录整理为标准药学监护记录。

输出格式（一行，不换行）：
问题：<患儿情况和当前用药>。分析：<药学分析>。处理：<建议措施>。结果/计划：<随访计划>。

规则：
1. 四段各以中文句号结尾，段落间无换行
2. 保留原始输入中的药物名称、剂量、检验数值，不虚构未提及的数据
3. 如某段信息不足，用简短通用表述补充
4. 语言专业简洁\
"""

_REQUIRED_MARKERS = ["问题：", "分析：", "处理：", "结果/计划："]


def _validate(text: str) -> bool:
    """Return True if all four section markers are present."""
    return all(m in text for m in _REQUIRED_MARKERS)


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

    user_content = ""
    if patient_info:
        user_content += f"患者信息：{patient_info}\n"
    user_content += f"查房记录：{raw_text}"

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

        # First attempt failed validation – retry with correction hint
        if attempt == 0:
            messages.append({"role": "assistant", "content": note})
            messages.append({
                "role": "user",
                "content": (
                    "格式不正确，请严格按照要求重新输出。"
                    '必须包含"问题："、"分析："、"处理："、"结果/计划："四个标记，'
                    "一行输出，不换行。"
                ),
            })

    # Both attempts failed validation – return what we have
    return {"note": note, "error": "格式校验未通过，请检查并手动修改。"}
