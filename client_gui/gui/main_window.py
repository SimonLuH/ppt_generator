import os
import json
import logging
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog,
    QMessageBox, QProgressBar, QPlainTextEdit,
    QFormLayout, QGroupBox, QSplitter
)
from PyQt5.QtCore import Qt
from client_gui.threads.worker_thread import WorkerThread
from client_gui.utils.logger import configure_logging
from client_gui.utils.exception_handler import show_error

logger = logging.getLogger(__name__)

class PPTClientGUI(QMainWindow):
    """
    主窗口：左侧表单 + 右侧日志(QPlainTextEdit)，
    支持自定义 slide_mappings.json，允许用户指定并行数量。
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.log_box = None
        self.input_max_workers = None
        self.label_max_workers = None
        self.progress_bar = None
        self.btn_stop = None
        self.btn_run = None
        self.btn_edit_mapping = None
        self.btn_select_mapping = None
        self.btn_output = None
        self.edit_output = None
        self.btn_excel = None
        self.edit_excel = None
        self.btn_template = None
        self.edit_template = None
        self.setWindowTitle("PPT生成客户端 - 可选自定义Mappings")
        self.resize(900, 500)
        self.worker = None
        self.mapping_file = None
        self.config_file = os.path.join(os.path.dirname(__file__), "gui_last_config.json")
        self.init_ui()
        self.load_settings()

    def init_ui(self):
        """
        初始化主界面，使用QSplitter分割左右区域。
        """
        # 输入控件
        self.edit_template = QLineEdit()
        self.btn_template = QPushButton("选择模板")
        self.edit_excel = QLineEdit()
        self.btn_excel = QPushButton("选择Excel目录")
        self.edit_output = QLineEdit()
        self.btn_output = QPushButton("选择输出目录")
        self.btn_select_mapping = QPushButton("选择Mappings文件")
        self.btn_edit_mapping = QPushButton("编辑slide_mappings")
        self.btn_run = QPushButton("开始处理")
        self.btn_stop = QPushButton("停止")
        self.progress_bar = QProgressBar()
        self.label_max_workers = QLabel("并行线程数:")
        self.input_max_workers = QLineEdit()
        self.input_max_workers.setPlaceholderText("默认: CPU核心数")

        # 表单布局
        form_layout = QFormLayout()
        form_layout.addRow("PPT模板:", self.edit_template)
        form_layout.addRow("", self.btn_template)
        form_layout.addRow("Excel目录:", self.edit_excel)
        form_layout.addRow("", self.btn_excel)
        form_layout.addRow("输出目录:", self.edit_output)
        form_layout.addRow("", self.btn_output)
        form_layout.addRow(self.label_max_workers, self.input_max_workers)
        group_box = QGroupBox("配置信息")
        group_box.setLayout(form_layout)

        # 按钮布局
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.btn_select_mapping)
        button_layout.addWidget(self.btn_edit_mapping)
        button_layout.addWidget(self.btn_run)
        button_layout.addWidget(self.btn_stop)

        # 左侧布局
        left_layout = QVBoxLayout()
        left_layout.addWidget(group_box)
        left_layout.addLayout(button_layout)
        left_layout.addWidget(self.progress_bar)
        left_widget = QWidget()
        left_widget.setLayout(left_layout)

        # 右侧日志
        self.log_box = QPlainTextEdit()
        self.log_box.setReadOnly(True)

        # 分割器
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(left_widget)
        splitter.addWidget(self.log_box)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)

        # 主布局
        container = QWidget()
        main_layout = QVBoxLayout()
        main_layout.addWidget(splitter)
        container.setLayout(main_layout)
        self.setCentralWidget(container)

        # 信号连接
        self.btn_template.clicked.connect(self.select_template)
        self.btn_excel.clicked.connect(self.select_excel_dir)
        self.btn_output.clicked.connect(self.select_output_dir)
        self.btn_select_mapping.clicked.connect(self.select_mapping_file)
        self.btn_edit_mapping.clicked.connect(self.edit_mappings)
        self.btn_run.clicked.connect(self.run_process)
        self.btn_stop.clicked.connect(self.stop_process)

    def load_settings(self):
        """
        加载上次的配置。
        """
        if os.path.isfile(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.edit_template.setText(data.get("template_path", ""))
                self.edit_excel.setText(data.get("excel_dir", ""))
                self.edit_output.setText(data.get("output_dir", ""))
                self.mapping_file = data.get("slide_mappings_file", None)
                if self.mapping_file:
                    self.log_box.appendPlainText(f"上次使用的映射文件: {self.mapping_file}")
                self.input_max_workers.setText(data.get("max_workers", ""))
            except Exception as e:
                logger.error(f"载入配置失败: {e}")

    def save_settings(self):
        """
        保存当前配置。
        """
        data = {
            "template_path": self.edit_template.text(),
            "excel_dir": self.edit_excel.text(),
            "output_dir": self.edit_output.text(),
            "slide_mappings_file": self.mapping_file,
            "max_workers": self.input_max_workers.text()
        }
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            logger.error(f"保存配置失败: {e}")

    def select_template(self):
        """
        选择PPT模板文件。
        """
        path, _ = QFileDialog.getOpenFileName(self, "选择PPT模板", ".", "PPTX文件(*.pptx)")
        if path:
            self.edit_template.setText(path)

    def select_excel_dir(self):
        """
        选择Excel目录。
        """
        directory = QFileDialog.getExistingDirectory(self, "选择Excel目录", ".")
        if directory:
            self.edit_excel.setText(directory)

    def select_output_dir(self):
        """
        选择输出目录。
        """
        directory = QFileDialog.getExistingDirectory(self, "选择输出目录", ".")
        if directory:
            self.edit_output.setText(directory)

    def select_mapping_file(self):
        """
        选择自定义的slide_mappings.json文件。
        """
        path, _ = QFileDialog.getOpenFileName(self, "选择Mappings文件", ".", "JSON文件(*.json)")
        if path:
            self.mapping_file = path
            self.log_box.appendPlainText(f"已选择映射文件: {path}")

    def edit_mappings(self):
        """
        打开映射编辑器窗口。
        """
        try:
            from client_gui.gui.slide_mapping_editor import SlideMappingEditor
            editor = SlideMappingEditor(self.mapping_file, self)
            if editor.exec_() == SlideMappingEditor.Accepted:
                self.mapping_file = editor.mapping_file
                self.log_box.appendPlainText(f"映射文件已更新: {self.mapping_file}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"编辑映射失败: {e}")

    def run_process(self):
        """
        开始处理PPT生成任务。
        """
        template_path = self.edit_template.text()
        excel_dir = self.edit_excel.text()
        output_dir = self.edit_output.text()

        # 验证输入
        if not os.path.isfile(template_path):
            QMessageBox.warning(self, "错误", "模板文件不存在！")
            return
        for path, desc in [(excel_dir, "Excel目录"), (output_dir, "输出目录")]:
            if not os.path.isdir(path):
                QMessageBox.warning(self, "错误", f"{desc}不存在！")
                return

        # 获取并行线程数
        max_workers_input = self.input_max_workers.text().strip()
        max_workers = int(max_workers_input) if max_workers_input.isdigit() else None
        if max_workers:
            self.log_box.appendPlainText(f"设置并行线程数为: {max_workers}")
        else:
            self.log_box.appendPlainText("未设置并行线程数，使用默认值 (CPU核心数)")

        self.log_box.appendPlainText("开始执行...")
        self.progress_bar.setValue(0)

        # 启动工作线程
        self.worker = WorkerThread(
            template_path=template_path,
            excel_dir=excel_dir,
            output_dir=output_dir,
            slide_mappings_file=self.mapping_file,
            max_workers=max_workers
        )
        self.worker.progress_update.connect(self.update_progress)
        self.worker.log_update.connect(self.update_log)
        self.worker.finished.connect(self.on_finished)
        self.worker.start()

    def stop_process(self):
        """
        停止处理任务（需实现停止逻辑）。
        """
        QMessageBox.information(self, "信息", "停止功能尚未实现。")

    def update_progress(self, value: int):
        """
        更新进度条。
        """
        self.progress_bar.setValue(value)

    def update_log(self, message: str):
        """
        更新日志显示。
        """
        self.log_box.appendPlainText(message)

    def on_finished(self):
        """
        处理完成后的操作。
        """
        self.log_box.appendPlainText("处理完成。")
        self.progress_bar.setValue(100)
        self.worker = None

    def closeEvent(self, event):
        """
        窗口关闭前保存配置。
        """
        self.save_settings()
        super().closeEvent(event)
