# PPT 生成器

## 概述

**PPT 生成器** 是一个基于 Python 的工具，旨在通过预定义的模板和动态数据源自动化创建 PowerPoint 演示文稿。该工具利用 PPT 模板中的占位符，并将其替换为来自 Excel 文件的数据，从而简化生成针对特定数据集的定制化演示文稿的过程。

## 功能特性

- **占位符替换：** 自动将 PPT 模板中的所有占位符替换为对应的 Excel 数据。
- **文本与表格处理：** 支持在所有幻灯片的文本框和表格中进行内容替换。
- **动态幻灯片复制：** 根据预定义的规则动态复制特定幻灯片，并填充相关数据。
- **批量处理：** 批量处理指定输入目录中的多个 Excel 文件，并在输出目录生成相应的 PPT 文件。
- **进度跟踪：** 显示实时进度条以监控处理状态。
- **可扩展架构：** 模块化的代码结构，便于维护和功能扩展。

## 目录

- [功能特性](#功能特性)
- [先决条件](#先决条件)
- [安装指南](#安装指南)
- [配置说明](#配置说明)
- [使用方法](#使用方法)
- [项目结构](#项目结构)
- [贡献指南](#贡献指南)
- [许可协议](#许可协议)
- [联系信息](#联系信息)

## 先决条件

- **Python 3.9 及以上版本**
- **pip**（Python 包管理器）

## 安装指南

1. **克隆仓库**

   ```bash
   git clone https://github.com/SimonLuH/ppt_generator.git
   cd ppt-generator
2. **创建虚拟环境**
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # Windows 用户使用: venv\Scripts\activate
3. 安装依赖包
   ```bash
   pip install -r requirements.txt

## 使用方法

1、**准备PPT模版**
- 使用 [占位符] 格式设计 PPT 模板中的占位符。
- 占位符可以存在于文本框或表格中。
2、**准备Excel数据源**
- 将数据组织在 Excel 文件（.xlsx）中，每个相关的工作表对应特定的幻灯片。
- 确保第一行包含与 PPT 模板中占位符匹配的标题。
3、**定义幻灯片映射关系**
- slide_mappings 文件示例
- 该映射确定每张幻灯片如何与数据表对应，以及是否需要复制幻灯片。
  ```python
  slide_mappings = {
    2:  {'type': 'row_for_page',      'sheet': '区域组织健康度构成'},
    3:  {'type': 'row_for_table_row', 'sheet': '区域高绩效人才盘点及目标规划'},
    # 根据需要添加更多映射
  }
4、**运行程序**
  ```python
  python gui_main.py
  ```
- 脚本将处理输入目录中的每个 Excel 文件，生成相应的 PPT 文件，并将其保存到输出目录。
- 处理过程中将显示进度条以显示当前状态。
