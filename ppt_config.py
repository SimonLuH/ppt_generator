"""
ppt_config.py
----------
存放全局配置信息的文件。例如：
- PPT 模板路径
- Excel 文件输入目录
- 输出 PPT 文件目录
如果需要更多自定义常量，也可放在此处。
"""

import os

#: PPT 模板文件的路径
PPT_TEMPLATE_PATH = r""

#: Excel 文件的输入目录
EXCEL_INPUT_DIR = r""

#: 生成 PPT 的输出目录
PPT_OUTPUT_DIR = r""

# 确保输出目录存在，否则创建
if not os.path.exists(PPT_OUTPUT_DIR):
    os.makedirs(PPT_OUTPUT_DIR, exist_ok=True)
