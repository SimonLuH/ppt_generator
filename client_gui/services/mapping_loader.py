import os
import json

# 默认映射，可根据需要在此处定义
default_mappings = {}

def load_slide_mappings(slide_mapping_config: str) -> dict:
    """
    从 slide_mappings.json 加载映射，若失败则返回 default_mappings
    并将所有 key 转换为 int 类型
    """
    if slide_mapping_config and os.path.isfile(slide_mapping_config):
        try:
            with open(slide_mapping_config, "r", encoding="utf-8") as f:
                raw_mappings = json.load(f)
            return _convert_keys_to_int(raw_mappings, slide_mapping_config)
        except Exception as e:
            print(f"读取映射失败: {e}, 使用默认.")
    else:
        print("未找到 slide_mappings.json, 采用默认写死映射.")
    return default_mappings

def _convert_keys_to_int(raw_mappings: dict, config_path: str) -> dict:
    """
    将字典的 key 尝试转换为 int，无法转换的跳过并打印警告
    """
    slide_mappings = {}
    for k, v in raw_mappings.items():
        try:
            slide_mappings[int(k)] = v
        except ValueError:
            print(f"Warning: 无法将 key='{k}' 转成 int, 跳过 (文件: {config_path}).")
    print(f"已从 {config_path} 载入 slide_mappings.")
    return slide_mappings