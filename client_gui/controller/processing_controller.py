import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Callable, Optional
import logging
import traceback  # 引入 traceback 模块以获取堆栈信息
from client_gui.services.excel_processor import process_excel_file
from client_gui.services.mapping_loader import load_slide_mappings

logger = logging.getLogger(__name__)

def run_processing(
    template_path: str,
    excel_dir: str,
    output_dir: str,
    slide_mappings_file: Optional[str] = None,
    max_workers: Optional[int] = None,
    progress_callback: Optional[Callable[[int], None]] = None,
    log_callback: Optional[Callable[[str], None]] = None
) -> None:
    """
    主处理逻辑：
    1. 收集Excel文件
    2. 加载映射配置
    3. 确保输出目录存在
    4. 并行处理Excel文件
    5. 更新进度和日志
    """

    logger.debug("开始运行 run_processing 函数。")
    logger.debug(f"输入参数 - template_path: {template_path}, excel_dir: {excel_dir}, "
                 f"output_dir: {output_dir}, slide_mappings_file: {slide_mappings_file}, "
                 f"max_workers: {max_workers}")

    # 验证模板文件路径
    if not os.path.isfile(template_path):
        msg = f"模板文件不存在: {template_path}"
        logger.error(msg)
        if log_callback:
            log_callback(msg)
        return
    else:
        logger.debug(f"找到模板文件: {template_path}")

    # 验证Excel目录路径
    if not os.path.isdir(excel_dir):
        msg = f"Excel目录不存在: {excel_dir}"
        logger.error(msg)
        if log_callback:
            log_callback(msg)
        return
    else:
        logger.debug(f"找到Excel目录: {excel_dir}")

    # 收集Excel文件
    try:
        excel_files = [f for f in os.listdir(excel_dir) if f.lower().endswith(".xlsx")]
        logger.debug(f"在目录 {excel_dir} 中找到 {len(excel_files)} 个xlsx文件: {excel_files}")
    except Exception as e:
        msg = f"无法读取Excel目录 {excel_dir}: {e}"
        logger.error(msg)
        logger.error(traceback.format_exc())  # 记录完整堆栈信息
        if log_callback:
            log_callback(msg + "\n" + traceback.format_exc())
        return

    total_files = len(excel_files)
    if total_files == 0:
        msg = f"在目录 {excel_dir} 中未找到任何xlsx文件。"
        logger.warning(msg)
        if log_callback:
            log_callback(msg)
        return

    # 加载映射配置
    try:
        slide_mapping = load_slide_mappings(slide_mappings_file)
        if not slide_mapping:
            logger.warning("加载的slide_mapping为空。")
        else:
            logger.debug(f"加载的slide_mapping内容: {slide_mapping}")
    except Exception as e:
        msg = f"加载slide_mappings文件时出错: {e}"
        logger.error(msg)
        logger.error(traceback.format_exc())  # 记录完整堆栈信息
        if log_callback:
            log_callback(msg + "\n" + traceback.format_exc())
        return

    # 确保输出目录存在
    if not os.path.isdir(output_dir):
        try:
            os.makedirs(output_dir, exist_ok=True)
            msg = f"已创建输出目录: {output_dir}"
            logger.info(msg)
            if log_callback:
                log_callback(msg)
        except Exception as e:
            msg = f"无法创建输出目录 {output_dir}: {e}"
            logger.error(msg)
            logger.error(traceback.format_exc())  # 记录完整堆栈信息
            if log_callback:
                log_callback(msg + "\n" + traceback.format_exc())
            return
    else:
        logger.debug(f"输出目录已存在: {output_dir}")

    # 确定并行线程数
    if max_workers and isinstance(max_workers, int):
        workers = max_workers
    else:
        workers = os.cpu_count() or 4
    msg = f"使用 {workers} 个并行线程处理Excel文件。"
    logger.info(msg)
    if log_callback:
        log_callback(msg)

    # 并行处理Excel文件
    processed = 0
    logger.debug("开始并行处理Excel文件。")
    try:
        with ThreadPoolExecutor(max_workers=workers) as executor:
            future_to_file = {
                executor.submit(
                    process_excel_file,
                    excel_file,
                    slide_mapping,
                    excel_dir,
                    output_dir,
                    template_path
                ): excel_file
                for excel_file in excel_files
            }
            logger.debug("所有Excel文件任务已提交。")

            for future in as_completed(future_to_file):
                excel_file = future_to_file[future]
                logger.debug(f"开始处理文件: {excel_file}")
                try:
                    result = future.result()
                    if result:
                        processed += 1
                        logger.debug(f"成功处理文件: {excel_file} ({processed}/{total_files})")
                    else:
                        logger.debug(f"跳过文件: {excel_file}")
                except Exception as e:
                    error_msg = f"处理 {excel_file} 时发生异常: {e}"
                    logger.error(error_msg)
                    logger.error(traceback.format_exc())  # 记录完整堆栈信息
                    if log_callback:
                        log_callback(error_msg + "\n" + traceback.format_exc())
                # 更新进度
                try:
                    percentage = int((processed / total_files) * 100)
                except ZeroDivisionError:
                    percentage = 0
                    logger.warning("总文件数为0，无法计算进度百分比。")
                if progress_callback:
                    progress_callback(percentage)
                logger.debug(f"当前进度: {percentage}%")

    except Exception as e:
        msg = f"多线程处理时发生异常: {e}"
        logger.error(msg)
        logger.error(traceback.format_exc())  # 记录完整堆栈信息
        if log_callback:
            log_callback(msg + "\n" + traceback.format_exc())
        return

    # 完成日志
    completion_msg = f"所有Excel处理完毕! 共处理 {processed} 个文件。输出目录: {output_dir}"
    logger.info(completion_msg)
    if log_callback:
        log_callback(completion_msg)
    logger.debug("run_processing 函数结束。")
