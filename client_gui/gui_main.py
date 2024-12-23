# gui_main.py

import sys, os, re, json
from typing import List
import logging
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox,
    QProgressBar, QPlainTextEdit, QFormLayout, QGroupBox, QSplitter
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed
from business_logic.processor import process_ppt_with_data
from data_access.excel_reader import ExcelDataProvider
import time

def resource_path(relative_path):
    """获取资源文件的绝对路径"""
    try:
        # PyInstaller 会将临时文件提取到 _MEIPASS 目录中
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# 定义处理逻辑
default_mappings = {}

def load_slide_mappings(slide_mapping_config):
    """
    从 slide_mappings.json 加载映射，若失败则返回 default_mappings
    并将所有 key 转换为 int 类型
    """
    if slide_mapping_config and os.path.isfile(slide_mapping_config):
        try:
            with open(slide_mapping_config, "r", encoding="utf-8") as f:
                raw_mappings = json.load(f)
            # 将 key 转为 int
            slide_mappings = {}
            for k, v in raw_mappings.items():
                try:
                    ik = int(k)
                    slide_mappings[ik] = v
                except ValueError:
                    print(f"Warning: 无法将 key='{k}' 转成int, 跳过.")
            print(f"已从 {slide_mapping_config} 载入 slide_mappings.")
            return slide_mappings
        except Exception as e:
            print(f"读取映射失败: {e}, 使用默认.")
    else:
        print("未找到 slide_mappings.json, 采用默认写死映射.")
    return default_mappings

def process_excel_file(excel_file, slide_mappings, input_dir, output_dir, template_path):
    """
    处理单个 Excel 文件，生成对应的 PPT
    """
    try:
        excel_path = os.path.join(input_dir, excel_file)
        base_name, _ = os.path.splitext(excel_file)
        output_ppt_filename = base_name + ".pptx"
        output_path = os.path.join(output_dir, output_ppt_filename)

        if os.path.exists(output_path):
            print(f"\n已存在同名 PPT，跳过: {output_ppt_filename}")
            return False  # 未处理

        # 这里需要您实现 ExcelDataProvider 和 process_ppt_with_data 的逻辑
        # 假设已正确导入相关模块并实现
        data_provider = ExcelDataProvider(excel_path)
        process_ppt_with_data(
            template_path=template_path,
            output_path=output_path,
            data_provider=data_provider,
            slide_mappings=slide_mappings
        )
        print(f"已处理: {excel_file} -> {output_ppt_filename}")
        return True  # 已处理
    except Exception as e:
        print(f"处理 {excel_file} 时出错: {e}")
        return False  # 处理失败

def run_main(template_path, excel_dir, output_dir, slide_mappings_file=None, max_workers=None, progress_callback=None, log_callback=None):
    """
    1) 从 EXCEL_INPUT_DIR 中收集 Excel
    2) 读取 slide_mappings.json (如失败则用默认映射)
    3) 使用多线程并行处理 Excel 文件
    4) 显示进度条
    """
    # 1) 收集 Excel 文件
    excel_files = [
        f for f in os.listdir(excel_dir)
        if f.lower().endswith(".xlsx")
    ]
    total_count = len(excel_files)
    if not excel_files:
        if log_callback:
            log_callback(f"在目录 {excel_dir} 中未找到任何 xlsx 文件.")
        print(f"在目录 {excel_dir} 中未找到任何 xlsx 文件.")
        return

    # 2) 加载 slide_mappings
    slide_mappings = load_slide_mappings(slide_mappings_file)

    # 3) 确保输出目录存在
    if not os.path.isdir(output_dir):
        try:
            os.makedirs(output_dir, exist_ok=True)
            if log_callback:
                log_callback(f"已创建输出目录: {output_dir}")
            print(f"已创建输出目录: {output_dir}")
        except Exception as e:
            if log_callback:
                log_callback(f"无法创建输出目录 {output_dir}: {e}")
            print(f"无法创建输出目录 {output_dir}: {e}")
            return

    # 4) 读取并指定 max_workers
    try:
        print(f"环境变量 MAX_WORKERS_ENV: {max_workers} (类型: {type(max_workers)})")  # 调试
        if max_workers and isinstance(max_workers, int):
            actual_max_workers = max_workers
        else:
            actual_max_workers = os.cpu_count()
    except Exception as e:
        print(f"解析 MAX_WORKERS 失败: {e}, 使用默认值 (CPU核心数).")
        actual_max_workers = os.cpu_count()

    print(f"使用 {actual_max_workers} 个并行线程处理 Excel 文件.")

    # 5) 使用 ThreadPoolExecutor 进行并行处理
    processed_count = 0
    try:
        with ThreadPoolExecutor(max_workers=actual_max_workers) as executor:
            # 提交所有任务
            future_to_file = {
                executor.submit(process_excel_file, excel_file, slide_mappings, excel_dir, output_dir, template_path): excel_file
                for excel_file in excel_files
            }

            for future in as_completed(future_to_file):
                excel_file = future_to_file[future]
                try:
                    result = future.result()
                    if result:
                        processed_count += 1
                except Exception as e:
                    print(f"处理 {excel_file} 时发生异常: {e}")
                    if log_callback:
                        log_callback(f"处理 {excel_file} 时发生异常: {e}")

                # 6) 更新进度条
                perc = processed_count / total_count
                if progress_callback:
                    progress_callback(int(perc * 100))
    except Exception as e:
        print(f"多线程处理时发生异常: {e}")
        if log_callback:
            log_callback(f"多线程处理时发生异常: {e}")
        return

    if log_callback:
        log_callback("所有Excel处理完毕!")
        log_callback(f"输出目录: {output_dir}")
    print("所有Excel处理完毕!")
    print(f"输出目录: {output_dir}")

class WorkerThread(QThread):
    """用于在后台运行 main.py 的工作线程"""
    progress_update = pyqtSignal(int)  # 信号，用于更新进度条
    log_update = pyqtSignal(str)       # 信号，用于更新日志
    finished = pyqtSignal()            # 信号，处理完成

    def __init__(self, template_path, excel_dir, output_dir, slide_mappings_file, max_workers):
        super().__init__()
        self.template_path = template_path
        self.excel_dir = excel_dir
        self.output_dir = output_dir
        self.slide_mappings_file = slide_mappings_file
        self.max_workers = max_workers

    def run(self):
        """在后台线程中运行 main.py 的逻辑"""
        try:
            # 调用 run_main 函数
            run_main(
                template_path=self.template_path,
                excel_dir=self.excel_dir,
                output_dir=self.output_dir,
                slide_mappings_file=self.slide_mappings_file,
                max_workers=self.max_workers,
                progress_callback=self.emit_progress,
                log_callback=self.emit_log
            )
            self.finished.emit()
        except Exception as e:
            self.log_update.emit(f"Exception: {e}")
            self.finished.emit()

    def emit_progress(self, value):
        """发出进度更新信号"""
        self.progress_update.emit(value)

    def emit_log(self, message):
        """发出日志更新信号"""
        self.log_update.emit(message)

class PPTClientGUI(QMainWindow):
    """
    主窗口: 左侧表单 + 右侧日志(QPlainTextEdit),
    支持"选择slide_mappings.json",将其传给 run_main
    并允许用户指定并行数量.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("PPT生成客户端 - 可选自定义Mappings")
        self.resize(900, 500)
        self.worker = None  # 用于存储工作线程
        self.config_file = os.path.join(os.path.dirname(__file__), "gui_last_config.json")
        self.mapping_file = None  # 用于存储用户指定的 slide_mappings.json
        self.initUI()
        self.loadSettings()

    def initUI(self):
        """初始化主界面(使用QSplitter分左右,左侧表单+进度,右侧日志)"""
        self.editTemplate = QLineEdit()
        self.btnTemplate = QPushButton("选择模板")
        self.editExcel = QLineEdit()
        self.btnExcel = QPushButton("选择Excel目录")
        self.editOutput = QLineEdit()
        self.btnOutput = QPushButton("选择输出目录")
        # 新增按钮 - 选择自定义slide_mappings.json
        self.btnSelectMappingFile = QPushButton("选择Mappings文件")
        self.btnEditMap = QPushButton("编辑slide_mappings")
        self.btnRun = QPushButton("开始处理")
        self.btnStop = QPushButton("停止")
        self.progress = QProgressBar()
        # 新增并行数量输入
        self.labelMaxWorkers = QLabel("并行线程数:")
        self.inputMaxWorkers = QLineEdit()
        self.inputMaxWorkers.setPlaceholderText("默认: CPU核心数")
        # 用QFormLayout放表单,便于美观整齐
        formLay = QFormLayout()
        formLay.addRow("PPT模板:", self.editTemplate)
        formLay.addRow("", self.btnTemplate)
        formLay.addRow("Excel目录:", self.editExcel)
        formLay.addRow("", self.btnExcel)
        formLay.addRow("输出目录:", self.editOutput)
        formLay.addRow("", self.btnOutput)
        formLay.addRow(self.labelMaxWorkers, self.inputMaxWorkers)  # 新增
        groupBox = QGroupBox("配置信息")
        groupBox.setLayout(formLay)
        # 下方按钮+进度用另一个布局
        btnLay = QHBoxLayout()
        btnLay.addWidget(self.btnSelectMappingFile)  # 新增
        btnLay.addWidget(self.btnEditMap)
        btnLay.addWidget(self.btnRun)
        btnLay.addWidget(self.btnStop)
        leftLay = QVBoxLayout()
        leftLay.addWidget(groupBox)
        leftLay.addLayout(btnLay)
        leftLay.addWidget(self.progress)
        leftWidget = QWidget()
        leftWidget.setLayout(leftLay)
        # 右侧日志用QPlainTextEdit,可多行滚动
        self.logBox = QPlainTextEdit()
        self.logBox.setReadOnly(True)
        rightWidget = self.logBox
        # 用QSplitter分割
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(leftWidget)
        splitter.addWidget(rightWidget)
        splitter.setStretchFactor(0, 1)  # 左面伸缩比
        splitter.setStretchFactor(1, 2)  # 右面伸缩比
        # 主布局
        container = QWidget()
        mainLay = QVBoxLayout()
        mainLay.addWidget(splitter)
        container.setLayout(mainLay)
        self.setCentralWidget(container)
        # 信号连接
        self.btnTemplate.clicked.connect(self.selectTemplate)
        self.btnExcel.clicked.connect(self.selectExcelDir)
        self.btnOutput.clicked.connect(self.selectOutputDir)
        self.btnSelectMappingFile.clicked.connect(self.selectMappingFile)  # 新增
        self.btnEditMap.clicked.connect(self.editMappings)
        self.btnRun.clicked.connect(self.runProcess)
        self.btnStop.clicked.connect(self.stopProcess)

    def loadSettings(self):
        """尝试从 gui_last_config.json 读取用户上次输入."""
        if os.path.isfile(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.editTemplate.setText(data.get("template_path", ""))
                self.editExcel.setText(data.get("excel_dir", ""))
                self.editOutput.setText(data.get("output_dir", ""))
                # 若有上次选择的映射文件,可记录
                self.mapping_file = data.get("slide_mappings_file", None)
                if self.mapping_file:
                    self.logBox.appendPlainText(f"上次使用的映射文件: {self.mapping_file}")
                # 加载并行数量
                self.inputMaxWorkers.setText(data.get("max_workers", ""))
            except Exception as e:
                print("载入配置失败:", e)

    def saveSettings(self):
        """保存当前输入到 gui_last_config.json,下次启动自动载入."""
        data = {
            "template_path": self.editTemplate.text(),
            "excel_dir": self.editExcel.text(),
            "output_dir": self.editOutput.text(),
            "slide_mappings_file": self.mapping_file,
            "max_workers": self.inputMaxWorkers.text()
        }
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print("保存配置失败:", e)

    def selectTemplate(self):
        """选择模板文件"""
        path, _ = QFileDialog.getOpenFileName(self, "选择PPT模板", ".", "PPTX文件(*.pptx)")
        if path:
            self.editTemplate.setText(path)

    def selectExcelDir(self):
        """选择Excel目录"""
        d = QFileDialog.getExistingDirectory(self, "选择Excel目录", ".")
        if d:
            self.editExcel.setText(d)

    def selectOutputDir(self):
        """选择输出目录"""
        d = QFileDialog.getExistingDirectory(self, "选择输出目录", ".")
        if d:
            self.editOutput.setText(d)

    def selectMappingFile(self):
        """让用户选自定义 slide_mappings.json 文件,后续执行将使用它"""
        path, _ = QFileDialog.getOpenFileName(self, "选择Mappings文件", ".", "JSON文件(*.json)")
        if path:
            self.mapping_file = path
            self.logBox.appendPlainText(f"已选择映射文件: {path}")

    def editMappings(self):
        """编辑 slide_mappings"""
        try:
            from client_gui.slide_mapping_editor import SlideMappingEditor
            dlg = SlideMappingEditor(self)
            dlg.exec_()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"编辑映射失败: {str(e)}")

    def runProcess(self):
        """直接调用 run_main 函数，并更新日志和进度"""
        if not os.path.isfile(self.editTemplate.text()):
            QMessageBox.warning(self, "错误", "模板不存在!")
            return
        for p in (self.editExcel.text(), self.editOutput.text()):
            if not os.path.isdir(p):
                QMessageBox.warning(self, "错误", f"目录不存在: {p}")
                return

        # 获取配置参数
        template_path = self.editTemplate.text()
        excel_dir = self.editExcel.text()
        output_dir = self.editOutput.text()
        slide_mappings_file = self.mapping_file if self.mapping_file and os.path.isfile(self.mapping_file) else None

        # 获取并行数量输入
        max_workers_input = self.inputMaxWorkers.text().strip()
        max_workers = int(max_workers_input) if max_workers_input.isdigit() else None
        if max_workers:
            self.logBox.appendPlainText(f"设置并行线程数为: {max_workers}")
        else:
            self.logBox.appendPlainText("未设置并行线程数, 使用默认值 (CPU核心数)")

        self.logBox.clear()
        self.logBox.appendPlainText("开始执行...")
        self.progress.setValue(0)

        # 创建并启动工作线程
        self.worker = WorkerThread(
            template_path=template_path,
            excel_dir=excel_dir,
            output_dir=output_dir,
            slide_mappings_file=slide_mappings_file,
            max_workers=max_workers
        )
        self.worker.progress_update.connect(self.updateProgress)
        self.worker.log_update.connect(self.updateLog)
        self.worker.finished.connect(self.onFinished)
        self.worker.start()

    def stopProcess(self):
        """停止进程"""
        # 由于使用 QThread，强制停止线程并不推荐。需要在 WorkerThread 中实现停止机制。
        QMessageBox.information(self, "信息", "停止功能尚未实现。")

    def updateProgress(self, value):
        """更新进度条"""
        self.progress.setValue(value)

    def updateLog(self, message):
        """更新日志"""
        self.logBox.appendPlainText(message)

    def onFinished(self):
        """进程结束"""
        self.logBox.appendPlainText("处理完成.")
        self.progress.setValue(100)
        self.worker = None

    def closeEvent(self, event):
        """窗口关闭前自动保存用户输入"""
        self.saveSettings()
        super().closeEvent(event)

def show_error(exc_type, exc_value, exc_traceback):
    """显示异常信息的弹窗并记录日志。"""
    error_message = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    logging.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)
    QMessageBox.critical(None, "程序出错", error_message)
    sys.exit(1)


def get_log_file_path():
    return os.path.join(os.path.expanduser("~"), "PPTClient_error.log")


# 配置日志记录
logging.basicConfig(
    filename=get_log_file_path(),
    level=logging.DEBUG,  # 设置为 DEBUG 以记录更多信息
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# 设置全局异常处理
sys.excepthook = show_error

def main():
    """GUI入口"""
    try:
        app = QApplication(sys.argv)
        w = PPTClientGUI()
        w.show()
        sys.exit(app.exec_())
    except Exception as e:
        raise RuntimeError(f"运行 run_main 时出错: {e}")

if __name__ == "__main__":
    main()
