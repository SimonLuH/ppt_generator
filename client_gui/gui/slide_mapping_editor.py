from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout,
    QPushButton, QPlainTextEdit, QMessageBox
)
import json
import os

class SlideMappingEditor(QDialog):
    """
    幻灯片映射编辑器对话框。
    """

    def __init__(self, mapping_file: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("编辑Slide Mappings")
        self.resize(600, 400)
        self.mapping_file = mapping_file
        self.init_ui()
        self.load_mappings()

    def init_ui(self):
        """
        初始化界面。
        """
        self.text_edit = QPlainTextEdit()
        self.btn_save = QPushButton("保存")
        self.btn_cancel = QPushButton("取消")

        button_layout = QHBoxLayout()
        button_layout.addStretch()
        button_layout.addWidget(self.btn_save)
        button_layout.addWidget(self.btn_cancel)

        main_layout = QVBoxLayout()
        main_layout.addWidget(self.text_edit)
        main_layout.addLayout(button_layout)
        self.setLayout(main_layout)

        # 信号连接
        self.btn_save.clicked.connect(self.save_mappings)
        self.btn_cancel.clicked.connect(self.reject)

    def load_mappings(self):
        """
        加载当前映射配置到文本编辑器。
        """
        if self.mapping_file and os.path.isfile(self.mapping_file):
            try:
                with open(self.mapping_file, "r", encoding="utf-8") as f:
                    content = f.read()
                self.text_edit.setPlainText(content)
            except Exception as e:
                QMessageBox.critical(self, "错误", f"加载映射文件失败: {e}")

    def save_mappings(self):
        """
        保存映射配置。
        """
        content = self.text_edit.toPlainText()
        try:
            mappings = json.loads(content)  # 验证JSON格式
            with open(self.mapping_file, "w", encoding="utf-8") as f:
                json.dump(mappings, f, indent=2, ensure_ascii=False)
            QMessageBox.information(self, "成功", "映射文件已保存。")
            self.accept()
        except json.JSONDecodeError as e:
            QMessageBox.warning(self, "错误", f"JSON格式错误: {e}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存映射文件失败: {e}")
