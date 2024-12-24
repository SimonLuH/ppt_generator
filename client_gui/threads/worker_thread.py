from PyQt5.QtCore import QThread, pyqtSignal
from client_gui.controller.processing_controller import run_processing

class WorkerThread(QThread):
    """
    后台工作线程，执行PPT生成任务。
    """
    progress_update = pyqtSignal(int)
    log_update = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(
        self,
        template_path: str,
        excel_dir: str,
        output_dir: str,
        slide_mappings_file: str,
        max_workers: int = None
    ):
        super().__init__()
        self.template_path = template_path
        self.excel_dir = excel_dir
        self.output_dir = output_dir
        self.slide_mappings_file = slide_mappings_file
        self.max_workers = max_workers

    def run(self):
        """
        执行PPT生成任务。
        """
        run_processing(
            template_path=self.template_path,
            excel_dir=self.excel_dir,
            output_dir=self.output_dir,
            slide_mappings_file=self.slide_mappings_file,
            max_workers=self.max_workers,
            progress_callback=self.emit_progress,
            log_callback=self.emit_log
        )
        self.finished.emit()

    def emit_progress(self, value: int):
        self.progress_update.emit(value)

    def emit_log(self, message: str):
        self.log_update.emit(message)
