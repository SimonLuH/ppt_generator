# PPT 生成客户端

## 概述

**PPT 生成客户端** 是一个基于 Python 的应用程序，旨在自动化地从 Excel 数据创建 PowerPoint 演示文稿。该工具利用 PyQt5 构建的用户友好图形界面，允许用户指定 PowerPoint 模板，选择多个 Excel 文件，并高效地生成定制化的演示文稿。它支持自定义幻灯片映射、多线程处理以提升性能，以及全面的日志记录以便监控和故障排除。

## 功能

- **自动化 PPT 生成**：根据预定义的模板将 Excel 数据转换为 PowerPoint 演示文稿。
- **自定义幻灯片映射**：定义 Excel 数据如何映射到 PPT 中的特定幻灯片和占位符。
- **用户友好的 GUI**：直观的界面，用于选择模板、Excel 目录、输出位置和映射配置。
- **多线程处理**：利用多个 CPU 核心并行处理 Excel 文件，减少处理时间。
- **进度跟踪**：实时进度条和日志窗口，用于监控 PPT 生成任务的状态。
- **错误处理与日志记录**：详细的日志记录成功操作和错误跟踪，帮助故障排除。
- **可配置设置**：保存和加载配置，简化重复任务。

## 目录

- [安装](#%E5%AE%89%E8%A3%85)
- [使用](#%E4%BD%BF%E7%94%A8)
    - [启动应用程序](#%E5%90%AF%E5%8A%A8%E5%BA%94%E7%94%A8%E7%A8%8B%E5%BA%8F)
    - [配置设置](#%E9%85%8D%E7%BD%AE%E8%AE%BE%E7%BD%AE)
    - [运行处理任务](#%E8%BF%90%E8%A1%8C%E5%A4%84%E7%90%86%E4%BB%BB%E5%8A%A1)
    - [编辑幻灯片映射](#%E7%BC%96%E8%BE%91%E5%B9%BB%E7%81%AF%E7%89%87%E6%98%A0%E5%B0%84)
- [项目结构](#%E9%A1%B9%E7%9B%AE%E7%BB%93%E6%9E%84)
- [依赖](#%E4%BE%9D%E8%B5%96)
- [贡献](#%E8%B4%A1%E7%8C%AE)
- [许可证](#%E8%AE%B8%E5%8F%AF%E8%AF%81)

## 安装

### 前提条件

- **Python 3.7+**：确保系统已安装 Python。可从 [python.org](https://www.python.org/downloads/) 下载。
- **pip**：Python 包管理器（通常随 Python 一起安装）。

### 克隆仓库

bash

复制代码

`git clone https://github.com/SimonLuH/ppt_generator.git`

### 创建虚拟环境（可选但推荐）

bash

复制代码

`python -m venv venv source venv/bin/activate  # Windows 系统使用: venv\Scripts\activate`

### 安装依赖

bash

复制代码

`pip install -r requirements.txt`

_注意：确保 `requirements.txt` 文件包含所有必要的包，如 `PyQt5`、`python-pptx`、`openpyxl` 等。_

## 使用

### 启动应用程序

运行主脚本以启动 GUI：

bash

复制代码

`python main.py`

### 配置设置

1. **选择 PPT 模板**：
    
    - 点击 **"选择模板"** 按钮。
    - 浏览并选择作为演示文稿模板的 `.pptx` 文件。
2. **选择 Excel 目录**：
    
    - 点击 **"选择Excel目录"** 按钮。
    - 选择包含要处理的 Excel (`.xlsx`) 文件的文件夹。
3. **选择输出目录**：
    
    - 点击 **"选择输出目录"** 按钮。
    - 指定生成的 PowerPoint 文件将保存的位置。
4. **配置并行线程数**：
    
    - 在 **"并行线程数"** 字段中输入要使用的并行线程数。如果留空，应用程序将默认使用 CPU 核心数。
5. **选择幻灯片映射文件**：
    
    - 点击 **"选择Mappings文件"** 按钮。
    - 选择自定义的 `slide_mappings.json` 文件，以定义 Excel 数据如何映射到 PPT 幻灯片和占位符。
6. **编辑幻灯片映射**：
    
    - 点击 **"编辑slide_mappings"** 按钮，打开幻灯片映射编辑器。
    - 修改 JSON 配置，以自定义数据插入到幻灯片中的方式。

### 运行处理任务

1. 配置完所有设置后，点击 **"开始处理"** 按钮，启动 PPT 生成过程。
2. **进度条** 将显示完成百分比。
3. **日志窗口** 将实时显示日志，包括成功操作和任何遇到的错误。

### 编辑幻灯片映射

1. 点击 **"编辑slide_mappings"** 按钮，打开幻灯片映射编辑器。
2. 编辑器将显示当前的 JSON 配置。根据需要修改，以定义 Excel 工作表和数据行如何对应到 PPT 模板中的特定幻灯片和占位符。
3. 点击 **"保存"** 应用更改，或点击 **"取消"** 放弃更改。

## 项目结构

```bash
ppt-generation-client/
│
├── business_logic/
│   └── processor.py          # 核心 PPT 生成处理逻辑
│
├── client_gui/
│   ├── controller/
│   │   └── processing_controller.py  # 处理任务控制器
│   ├── gui/
│   │   └── main_window.py    # 主 GUI 窗口【应用程序入口】
│   ├── model/
│   │   └── mapping_model.py  # 映射数据模型
│   ├── services/
│   │   ├── excel_processor.py       # 处理单个 Excel 文件
│   │   └── mapping_loader.py        # 加载幻灯片映射配置
│   ├── threads/
│   │   └── worker_thread.py   # 后台任务线程
│   └── utils/
│       ├── exception_handler.py       # 全局异常处理
│       ├── logger.py                 # 日志配置
│       └── resources.py              # 资源管理
│
├── data_access/
│   ├── base_provider.py       # 数据提供者抽象基类
│   └── excel_reader.py        # 从 Excel 文件读取数据
│
├── ppt_engine/
│   ├── deck_manager.py        # PPT Deck 操作管理
│   ├── placeholders.py        # 占位符替换逻辑
│   └── slide_handler.py       # 幻灯片特定操作处理
│
├── utils/
│   ├── configure_logging.py   # 日志设置
│   └── exception_handler.py   # 异常处理工具
│
├── gui_last_config.json      # 存储上次使用的配置
├── requirements.txt           # Python 依赖包                  
```

## 依赖

项目依赖以下 Python 包：

- **PyQt5**：用于构建图形用户界面。
- **python-pptx**：用于操作 PowerPoint 文件。
- **openpyxl**：用于读取 Excel 文件。
- **concurrent.futures**：用于多线程处理。
- **logging**：用于日志记录事件和错误。

确保通过 `pip install -r requirements.txt` 安装所有依赖。

## 贡献

欢迎贡献！如果您遇到任何问题或有改进建议，请提交问题或拉取请求。

1. Fork 仓库。
2. 创建新分支：`git checkout -b feature/YourFeature`。
3. 进行更改并提交：`git commit -m '添加新功能'`。
4. 推送到分支：`git push origin feature/YourFeature`。
5. 打开拉取请求。

## 许可证

本项目遵循 MIT 许可证。

---

_如有任何问题或需要支持，请联系 your.email@example.com。_
