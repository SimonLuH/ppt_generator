import os
import logging
from business_logic.processor import process_ppt_with_data
from client_gui.model.mapping_model import SlideMapping
from data_access.excel_reader import ExcelDataProvider

logger = logging.getLogger(__name__)

def process_excel_file(
    excel_file: str,
    slide_mapping,
    input_dir: str,
    output_dir: str,
    template_path: str
) -> bool:
    """
    处理单个Excel文件，生成对应的PPT。
    """
    try:
        excel_path = os.path.join(input_dir, excel_file)
        base_name, _ = os.path.splitext(excel_file)
        output_ppt_filename = f"{base_name}.pptx"
        output_path = os.path.join(output_dir, output_ppt_filename)

        if os.path.exists(output_path):
            logger.info(f"已存在同名PPT，跳过: {output_ppt_filename}")
            return False  # 未处理

        provider = ExcelDataProvider(excel_path)
        process_ppt_with_data(
            template_path=template_path,
            output_path=output_path,
            data_provider=provider,
            slide_mappings=slide_mapping
        )
        logger.info(f"已处理: {excel_file} -> {output_ppt_filename}")
        return True  # 已处理
    except Exception as e:
        logger.error(f"处理 {excel_file} 时出错: {e}")
        return False  # 处理失败
