import sys
import os

def resource_path(relative_path: str) -> str:
    """
    获取资源文件的绝对路径，兼容 PyInstaller。
    """
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_log_file_path() -> str:
    """
    获取日志文件路径。
    """
    return os.path.join(os.path.expanduser("~"), "PPTClient_error.log")
