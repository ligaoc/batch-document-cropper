"""文件格式验证器"""
import os
from typing import Tuple

# 支持的文件格式 (小写)
SUPPORTED_FORMATS = {".pdf", ".docx", ".doc"}


def get_file_extension(file_path: str) -> str:
    """获取文件扩展名 (小写)
    
    Args:
        file_path: 文件路径
        
    Returns:
        小写的文件扩展名 (包含点号)
    """
    _, ext = os.path.splitext(file_path)
    return ext.lower()


def is_supported_format(file_path: str) -> bool:
    """判断文件格式是否支持
    
    Args:
        file_path: 文件路径
        
    Returns:
        是否为支持的格式
    """
    ext = get_file_extension(file_path)
    return ext in SUPPORTED_FORMATS


def validate_file(file_path: str) -> Tuple[bool, str]:
    """验证文件是否可处理
    
    Args:
        file_path: 文件路径
        
    Returns:
        (是否有效, 错误信息)
    """
    if not file_path:
        return False, "文件路径不能为空"
    
    ext = get_file_extension(file_path)
    if not ext:
        return False, "文件没有扩展名"
    
    if not is_supported_format(file_path):
        supported = ", ".join(sorted(SUPPORTED_FORMATS))
        return False, f"不支持的文件格式: {ext}。支持的格式: {supported}"
    
    return True, ""


def is_pdf(file_path: str) -> bool:
    """判断是否为 PDF 文件"""
    return get_file_extension(file_path) == ".pdf"


def is_docx(file_path: str) -> bool:
    """判断是否为 DOCX 文件"""
    return get_file_extension(file_path) == ".docx"


def is_doc(file_path: str) -> bool:
    """判断是否为 DOC 文件"""
    return get_file_extension(file_path) == ".doc"


def needs_conversion(file_path: str) -> bool:
    """判断文件是否需要转换为 PDF
    
    Args:
        file_path: 文件路径
        
    Returns:
        是否需要转换 (DOC/DOCX 返回 True)
    """
    ext = get_file_extension(file_path)
    return ext in {".doc", ".docx"}


def needs_doc_to_docx_conversion(file_path: str) -> bool:
    """判断文件是否需要从 DOC 转换为 DOCX
    
    Args:
        file_path: 文件路径
        
    Returns:
        是否需要转换 (DOC 返回 True)
    """
    return is_doc(file_path)


def get_output_extension(file_path: str) -> str:
    """根据输入文件类型获取输出文件扩展名
    
    Args:
        file_path: 输入文件路径
        
    Returns:
        输出文件扩展名 (包含点号)
        - PDF -> .pdf
        - DOCX -> .docx
        - DOC -> .docx (转换后)
    """
    ext = get_file_extension(file_path)
    if ext == ".pdf":
        return ".pdf"
    elif ext in {".docx", ".doc"}:
        return ".docx"
    return ext
