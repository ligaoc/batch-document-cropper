"""输出文件管理模块"""
import os
from typing import Tuple

from .file_validator import get_output_extension


class OutputError(Exception):
    """输出错误"""
    def __init__(self, output_path: str, reason: str):
        self.output_path = output_path
        self.reason = reason
        super().__init__(f"输出失败 {output_path}: {reason}")


def generate_output_filename(input_path: str, suffix: str = "_cropped") -> str:
    """生成输出文件名（保持原始格式）
    
    Args:
        input_path: 输入文件路径
        suffix: 文件名后缀
        
    Returns:
        输出文件名 (不含目录)
        - PDF 输入 -> xxx_cropped.pdf
        - DOCX 输入 -> xxx_cropped.docx
        - DOC 输入 -> xxx_cropped.docx (转换后)
    """
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    output_ext = get_output_extension(input_path)
    return f"{base_name}{suffix}{output_ext}"


def generate_output_path(input_path: str, output_dir: str,
                         suffix: str = "_cropped") -> str:
    """生成完整输出路径（保持原始格式）
    
    Args:
        input_path: 输入文件路径
        output_dir: 输出目录
        suffix: 文件名后缀
        
    Returns:
        完整输出路径
    """
    filename = generate_output_filename(input_path, suffix)
    return os.path.join(output_dir, filename)


def validate_output_dir(output_dir: str) -> Tuple[bool, str]:
    """验证输出目录
    
    Args:
        output_dir: 输出目录路径
        
    Returns:
        (是否有效, 错误信息)
    """
    if not output_dir:
        return False, "输出目录不能为空"
    
    # 如果目录不存在，尝试创建
    if not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir, exist_ok=True)
        except PermissionError:
            return False, "没有权限创建输出目录"
        except Exception as e:
            return False, f"无法创建输出目录: {e}"
    
    # 检查是否为目录
    if not os.path.isdir(output_dir):
        return False, "输出路径不是目录"
    
    # 检查写入权限
    if not os.access(output_dir, os.W_OK):
        return False, "没有写入权限"
    
    return True, ""


def ensure_output_dir(output_dir: str) -> None:
    """确保输出目录存在
    
    Args:
        output_dir: 输出目录路径
        
    Raises:
        OutputError: 无法创建目录时抛出
    """
    valid, error = validate_output_dir(output_dir)
    if not valid:
        raise OutputError(output_dir, error)
