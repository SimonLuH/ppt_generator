# business_logic/processor.py

from typing import Dict, List, Any
from ppt_engine.deck_manager import open_ppt, save_ppt, close_ppt, copy_slide_after
from ppt_engine.slide_handler import fill_table_with_rows, fill_table_with_single_dict

def process_ppt_with_data(template_path: str, output_path: str, data_provider,
                          slide_mappings: Dict[int, Dict[str, Any]]):
    """
    两阶段：
      1) prepare_slides -> 先复制所有需要多份的幻灯片
      2) fill_placeholders -> 再统一占位符替换
    """
    # 1. 打开PPT模板
    prs = open_ppt(template_path)

    # 2. 读取Excel数据
    all_data = data_provider.read_data()

    # 3. 幻灯片布局(复制)
    fill_plan = prepare_slides(prs, slide_mappings, all_data)

    # 4. 填充占位符
    fill_placeholders(prs, fill_plan, all_data)

    # 5. 保存&关闭
    save_ppt(prs, output_path)

def prepare_slides(prs, slide_mappings: Dict[int, dict], all_data: dict) -> List[dict]:
    """
    跟原先一样: 根据 slide_mappings 先复制需要多份的幻灯片
    并返回 fill_plan
    """
    fill_plan = []
    sorted_keys = sorted(slide_mappings.keys())

    # offset_map 维护复制时的索引偏移
    offset_map = {k: k for k in sorted_keys}

    # python-pptx slides[] 是 0-based,
    # 但你原先映射中 slide_idx 是 1-based.
    # => 我们会在 fill_placeholders 里做 "slides[idx-1]"
    for k in sorted_keys:
        cfg = slide_mappings[k]
        sheet_name = cfg.get("sheet")
        data_type  = cfg.get("type")
        do_copy    = cfg.get("copy", False)
        data_rows  = all_data.get(sheet_name, [])
        n_rows     = len(data_rows)

        real_idx = offset_map[k]

        if not do_copy or n_rows <= 1:
            # 不复制 or 只有1行 => 只用一张
            fill_plan.append({
                "slide_index": real_idx,
                "sheet_name": sheet_name,
                "type": data_type,
                "row_data_index": 0,
                "copy_mode": False
            })
        else:
            # 需要复制 => n_rows多张
            copies_needed = n_rows - 1
            copy_slide_after(prs, base_index=real_idx, count=copies_needed)
            # 更新 offset_map
            for other_k in sorted_keys:
                if other_k >= k and other_k != k:
                    offset_map[other_k] += copies_needed
            # 记录 fill_plan
            for i in range(n_rows):
                fill_plan.append({
                    "slide_index": real_idx + i,
                    "sheet_name": sheet_name,
                    "type": data_type,
                    "row_data_index": i,
                    "copy_mode": True
                })

    return fill_plan

def fill_placeholders(prs, fill_plan: List[dict], all_data: dict):
    """
    复制完后, 幻灯片数量和顺序已固定
    我们遍历 fill_plan,
    对 slide_index 那张幻灯片做替换
    """
    fill_plan_sorted = sorted(fill_plan, key=lambda x: x["slide_index"])

    for item in fill_plan_sorted:
        idx        = item["slide_index"]  # 1-based
        sheet_name = item["sheet_name"]
        data_type  = item["type"]
        row_i      = item["row_data_index"]

        data_rows  = all_data.get(sheet_name, [])

        # python-pptx slides是0-based => slides[idx-1]
        if (idx-1) < 0 or (idx-1) >= len(prs.slides):
            print(f"[fill_placeholders] 幻灯片索引{idx}超范围, 跳过.")
            continue

        slide = prs.slides[idx - 1]
        if data_type == "row_for_table_row":
            # => 多行 => 同一张
            fill_table_with_rows(slide, data_rows)
        else:
            # => 一行 => 整张
            row_data = data_rows[row_i] if 0 <= row_i < len(data_rows) else {}
            fill_table_with_single_dict(slide, row_data)