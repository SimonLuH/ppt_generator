from typing import Dict

class SlideMapping:
    """
    表示幻灯片映射配置。
    """
    def __init__(self, mappings: Dict[int, dict]):
        self.mappings = mappings

    @staticmethod
    def from_dict(raw_mappings: Dict[str, dict]) -> 'SlideMapping':
        """
        从字典创建 SlideMapping 实例，键转换为 int。
        """
        converted = {}
        for k, v in raw_mappings.items():
            try:
                key_int = int(k)
                converted[key_int] = v
            except ValueError:
                logging.warning(f"无法将 key='{k}' 转成 int，跳过。")
        logging.info("Slide mappings loaded successfully.")
        return SlideMapping(converted)
