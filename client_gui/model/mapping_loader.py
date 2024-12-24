import os
import json
from typing import Optional
from client_gui.model.mapping_model import SlideMapping
from client_gui.utils import configure_logging

configure_logging()
logger = logging.getLogger(__name__)

DEFAULT_MAPPINGS = {}

def load_slide_mappings(config_path: Optional[str]) -> dict:
    """
    加载幻灯片映射配置，如果失败则返回默认映射。
    """
    if config_path and os.path.isfile(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                raw_mappings = json.load(f)
            logger.info(f"从 {config_path} 载入 slide_mappings.")
            return raw_mappings
        except Exception as e:
            logger.error(f"读取映射失败: {e}，使用默认映射。")
    else:
        logger.warning("未找到 slide_mappings.json，使用默认映射。")
    return DEFAULT_MAPPINGS
