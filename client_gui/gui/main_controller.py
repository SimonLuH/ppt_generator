import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Callable, Optional
from client_gui.services.mapping_loader import load_slide_mappings
from client_gui.services.excel_processor import process_excel_file

def run_main(
    template_path: str,
    excel_dir: str,
    output_dir: str,
    slide_mappings_file: Optional[str] = None,
    max_workers: Optional[int] = None,
    progress_callback: Optional[Callable[[int], None]] = None,
    log_callback: Optional[Callable[[str], None]] = None
) -> None:
    """
    业务主入口：
    1) 收集Excel文件并加载映射
    2) 按并行线程数处理Excel
    3) 进度回调和日志回调
    """
    excel_files = _collect_excel_files(excel_dir, log_callback)
    if not excel_files:
        return

    slide_mappings = load_slide_mappings(slide_mappings_file)
    _ensure_output_dir(output_dir, log_callback)
    actual_workers = _choose_worker_count(max_workers, log_callback)
    _process_all_excels(
        excel_files, slide_mappings, excel_dir, output_dir,
        template_path, actual_workers, progress_callback, log_callback
    )
    _finish_log(output_dir, log_callback)

def _collect_excel_files(excel_dir: str, log_callback: Optional[Callable[[str], None]]) -> list:
    """
    收集目录下所有 xlsx 文件
    """
    excel_files = [f for f in os.listdir(excel_dir) if f.lower().endswith(".xlsx")]
    if not excel_files:
        msg = f"在目录 {excel_dir} 中未找到任何 xlsx 文件."
        if log_callback:
            log_callback(msg)
        print(msg)
    return excel_files

def _ensure_output_dir(output_dir: str, log_callback: Optional[Callable[[str], None]]) -> None:
    """
    确保输出目录存在
    """
    if not os.path.isdir(output_dir):
        try:
            os.makedirs(output_dir, exist_ok=True)
            msg = f"已创建输出目录: {output_dir}"
            if log_callback:
                log_callback(msg)
            print(msg)
        except Exception as e:
            msg = f"无法创建输出目录 {output_dir}: {e}"
            if log_callback:
                log_callback(msg)
            print(msg)

def _choose_worker_count(max_workers: Optional[int], log_callback: Optional[Callable[[str], None]]) -> int:
    """
    解析可用线程数
    """
    try:
        if max_workers and isinstance(max_workers, int):
            return max_workers
        return os.cpu_count() or 4
    except Exception as e:
        msg = f"解析 MAX_WORKERS 失败: {e}, 使用默认值 (CPU核心数)."
        print(msg)
        if log_callback:
            log_callback(msg)
        return os.cpu_count() or 4

def _process_all_excels(
    excel_files: list,
    slide_mappings: dict,
    excel_dir: str,
    output_dir: str,
    template_path: str,
    max_workers: int,
    progress_callback: Optional[Callable[[int], None]],
    log_callback: Optional[Callable[[str], None]]
) -> None:
    """
    并行处理所有 Excel 文件
    """
    total_count = len(excel_files)
    processed_count = 0
    msg = f"使用 {max_workers} 个并行线程处理 Excel 文件."
    print(msg)
    if log_callback:
        log_callback(msg)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_file = {
            executor.submit(
                process_excel_file,
                excel_file, slide_mappings, excel_dir, output_dir, template_path
            ): excel_file
            for excel_file in excel_files
        }
        for future in as_completed(future_to_file):
            excel_file = future_to_file[future]
            try:
                if future.result():
                    processed_count += 1
            except Exception as e:
                err = f"处理 {excel_file} 时发生异常: {e}"
                print(err)
                if log_callback:
                    log_callback(err)
            _update_progress(processed_count, total_count, progress_callback)

def _update_progress(
    processed_count: int,
    total_count: int,
    progress_callback: Optional[Callable[[int], None]]
) -> None:
    """
    更新进度条
    """
    if progress_callback:
        percentage = int((processed_count / total_count) * 100)
        progress_callback(percentage)

def _finish_log(output_dir: str, log_callback: Optional[Callable[[str], None]]) -> None:
    """
    处理结束后输出日志
    """
    msg1 = "所有Excel处理完毕!"
    msg2 = f"输出目录: {output_dir}"
    print(msg1)
    print(msg2)
    if log_callback:
        log_callback(msg1)
        log_callback(msg2)