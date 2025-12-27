"""DOCX 裁剪器 - 通过调整页面边距实现裁剪"""
import os
from typing import Optional

from docx import Document
from docx.shared import Mm

from ..models.margin_settings import MarginSettings
from ..models.task import CropResult


class DOCXCropError(Exception):
    """DOCX 裁剪错误"""
    def __init__(self, file_path: str, reason: str):
        self.file_path = file_path
        self.reason = reason
        super().__init__(f"DOCX 裁剪失败 {file_path}: {reason}")


class DOCXCropper:
    """DOCX 裁剪器，通过调整页面边距实现裁剪"""
    
    def __init__(self):
        """初始化裁剪器"""
        pass
    
    def crop(self, input_path: str, output_path: str,
             margins: MarginSettings) -> CropResult:
        """裁剪 DOCX 文件（通过调整页面边距）
        
        Args:
            input_path: 输入 DOCX 路径
            output_path: 输出 DOCX 路径
            margins: 边距设置
            
        Returns:
            裁剪结果对象
        """
        try:
            # 验证输入文件
            if not os.path.isfile(input_path):
                return CropResult.failure_result(input_path, "输入文件不存在")
            
            # 验证边距设置
            is_valid, error_msg = margins.validate()
            if not is_valid:
                return CropResult.failure_result(input_path, error_msg)
            
            # 打开文档
            doc = Document(input_path)
            
            # 获取节数
            section_count = len(doc.sections)
            
            # 调整每个节的页面边距
            for section in doc.sections:
                section.top_margin = Mm(margins.top)
                section.bottom_margin = Mm(margins.bottom)
                section.left_margin = Mm(margins.left)
                section.right_margin = Mm(margins.right)
            
            # 确保输出目录存在
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            
            # 保存文档
            doc.save(output_path)
            
            return CropResult.success_result(
                input_path=input_path,
                output_path=output_path,
                pages=section_count,  # 使用节数作为页数估算
                orig_res=0,  # DOCX 没有分辨率概念
                out_res=0,
            )
            
        except Exception as e:
            return CropResult.failure_result(input_path, str(e))
    
    def get_section_count(self, input_path: str) -> int:
        """获取 DOCX 节数
        
        Args:
            input_path: DOCX 文件路径
            
        Returns:
            节数
        """
        try:
            doc = Document(input_path)
            return len(doc.sections)
        except Exception:
            return 0
    
    def get_current_margins(self, input_path: str) -> Optional[MarginSettings]:
        """获取 DOCX 当前边距设置（第一节）
        
        Args:
            input_path: DOCX 文件路径
            
        Returns:
            边距设置，失败返回 None
        """
        try:
            doc = Document(input_path)
            if not doc.sections:
                return None
            
            section = doc.sections[0]
            return MarginSettings(
                top=section.top_margin.mm if section.top_margin else 0,
                bottom=section.bottom_margin.mm if section.bottom_margin else 0,
                left=section.left_margin.mm if section.left_margin else 0,
                right=section.right_margin.mm if section.right_margin else 0,
            )
        except Exception:
            return None
