# ppt_engine/placeholders.py

import re

BRACKET_PATTERN = re.compile(r"\[[^\]]*\]")


def replace_placeholders(tf, row_data: dict):
    """
        在 tf(paragraphs/runs) 内做占位符替换, 保留原 run 样式.
        bracket_pattern 用于匹配形如 [xxx].
        row_data 是 { "[A]":"valA", "[B]":"valB", ... }

        不返回任何内容, 因为替换直接改 tf 内的 run.text.
        """
    for paragraph in tf.paragraphs:
        for run in paragraph.runs:
            old_run_text = run.text
            if not old_run_text:
                continue

            # 查找占位符
            matches = BRACKET_PATTERN.findall(old_run_text)
            if not matches:
                continue

            new_run_text = old_run_text
            for ph in matches:
                repl_val = str(row_data.get(ph, "未知"))
                new_run_text = new_run_text.replace(ph, repl_val)

            if new_run_text != old_run_text:
                run.text = new_run_text