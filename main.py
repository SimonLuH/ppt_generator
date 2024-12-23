# main.py

import os
import json
from concurrent.futures import ProcessPoolExecutor, as_completed
from data_access.excel_reader import ExcelDataProvider
from business_logic.processor import process_ppt_with_data

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
    3) 使用多进程并行处理 Excel 文件
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

    print(f"使用 {actual_max_workers} 个并行进程处理 Excel 文件.")

    # 5) 使用 ProcessPoolExecutor 进行并行处理
    processed_count = 0
    try:
        with ProcessPoolExecutor(max_workers=actual_max_workers) as executor:
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
                # Optionally, log progress
    except Exception as e:
        print(f"多进程处理时发生异常: {e}")
        if log_callback:
            log_callback(f"多进程处理时发生异常: {e}")
        return

    if log_callback:
        log_callback("所有Excel处理完毕!")
        log_callback(f"输出目录: {output_dir}")
    print("所有Excel处理完毕!")
    print(f"输出目录: {output_dir}")

if __name__ == "__main__":
    # 仅在作为脚本运行时执行
    run_main(
        template_path=os.environ.get("PPT_TEMPLATE_PATH", ""),
        excel_dir=os.environ.get("EXCEL_INPUT_DIR", ""),
        output_dir=os.environ.get("PPT_OUTPUT_DIR", ""),
        slide_mappings_file=os.environ.get("SLIDE_MAPPINGS_FILE", ""),
        max_workers=int(os.environ.get("MAX_WORKERS", "20"))
    )
