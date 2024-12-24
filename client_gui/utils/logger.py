import logging
import os
from client_gui.utils.resources import get_log_file_path

def configure_logging() -> None:
    """
    配置全局日志
    """
    logging.basicConfig(
        filename=get_log_file_path(),
        level=logging.DEBUG,  # 记录更多信息
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
