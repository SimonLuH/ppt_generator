# ppt_engine/slide_handler.py

from typing import List, Dict, Any
from ppt_engine.placeholders import replace_placeholders

def fill_table_with_rows(slide, data_rows: List[Dict[str, Any]]):
    """
    多行数据 -> 同一张表格:
    - 第1行是表头, 从第2行起写 data_rows
    - 若 data_rows 超过表格行数, 只写到最后
    """
    if not data_rows:
        return

    for shape in slide.shapes:
        if shape.has_table:  # python-pptx 判断表格
            table = shape.table
            row_count = table.rows.__len__()
            col_count = table.columns.__len__()

            write_start = 1  # 第1行(索引0)作为表头 => 从 row=1开始
            for i, row_data in enumerate(data_rows):
                current_row = write_start + i
                if current_row >= row_count:
                    break
                for c in range(col_count):
                    cell = table.cell(current_row, c)
                    if cell.text_frame:
                        replace_placeholders(cell.text_frame, row_data)

        elif shape.has_text_frame:  # 如果是文本框
            replace_placeholders(shape.text_frame, data_rows)

def fill_table_with_single_dict(slide, row_data: Dict[str, Any]):
    """
    一行数据 -> 整个表格(不做多行循环).
    也遍历文本框, 用 row_data 替换占位符
    """
    if not row_data:
        return

    for shape in slide.shapes:
        if shape.has_table:
            table = shape.table
            row_count = table.rows.__len__()
            col_count = table.columns.__len__()

            # 遍历所有行列
            for r in range(row_count):
                for c in range(col_count):
                    cell = table.cell(r, c)
                    if cell.text_frame:
                        replace_placeholders(cell.text_frame, row_data)

        elif shape.has_text_frame:
            replace_placeholders(shape.text_frame, row_data)