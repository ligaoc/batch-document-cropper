"""DOCX 裁剪器 - PDF 真正裁剪 + DOCX 调整边距"""
import os
import tempfile
import shutil
import logging
import time
from typing import Optional, Tuple

from ..models.margin_settings import MarginSettings
from ..models.task import CropResult
from .document_converter import get_document_converter
from .pdf_cropper import PDFCropper

logger = logging.getLogger(__name__)


class DOCXCropError(Exception):
    """DOCX 裁剪错误"""
    def __init__(self, file_path: str, reason: str):
        self.file_path = file_path
        self.reason = reason
        super().__init__(f"DOCX 裁剪失败 {file_path}: {reason}")


class DOCXCropper:
    """DOCX 裁剪器
    
    裁剪流程：
    1. DOCX → PDF（使用 WPS/Word）→ 裁剪 PDF（真正裁剪内容）
    2. DOCX → 调整边距的 DOCX（保留原始格式，文字可编辑）
    
    输出：
    - _cropped.pdf: 真正裁剪后的 PDF（内容被裁剪）
    - _cropped.docx: 调整边距后的 DOCX（保留原始格式，可编辑）
    """
    
    def __init__(self):
        """初始化裁剪器"""
        self._converter = get_document_converter()
        self._pdf_cropper = PDFCropper()
        print(f"[DOCX裁剪器] 初始化完成，转换器: {self._converter.converter_name}")
    
    def crop(self, input_path: str, output_path: str,
             margins: MarginSettings) -> CropResult:
        """裁剪 DOCX 文件
        
        会同时输出 PDF 和 DOCX 两种格式：
        - output_path: DOCX 文件（调整边距版，保留原始格式）
        - output_path 同目录下的 .pdf 文件（真正裁剪版）
        
        Args:
            input_path: 输入 DOCX 路径
            output_path: 输出 DOCX 路径
            margins: 裁剪量设置（从页面边缘裁剪掉的毫米数）
            
        Returns:
            裁剪结果对象
        """
        temp_dir = None
        temp_pdf = None
        
        try:
            # 验证输入文件
            if not os.path.isfile(input_path):
                return CropResult.failure_result(input_path, "输入文件不存在")
            
            # 验证裁剪量设置
            is_valid, error_msg = margins.validate()
            if not is_valid:
                return CropResult.failure_result(input_path, error_msg)
            
            # 检查转换器是否可用
            if not self._converter.is_available:
                return CropResult.failure_result(
                    input_path, 
                    "没有可用的转换器（需要 Microsoft Word 或 WPS）"
                )
            
            # 创建临时目录
            temp_dir = tempfile.mkdtemp()
            
            # 确定输出路径
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            
            # PDF 输出路径（与 DOCX 同名但扩展名为 .pdf）
            base_name = os.path.splitext(output_path)[0]
            pdf_output_path = base_name + ".pdf"
            
            # 步骤 1: DOCX → PDF → 裁剪 PDF
            print(f"[DOCX裁剪] 步骤1: 转换为 PDF 并裁剪")
            logger.info(f"DOCX → PDF 转换，使用: {self._converter.converter_name}")
            temp_pdf = self._converter.convert_to_pdf(input_path, temp_dir)
            
            # 裁剪 PDF
            crop_result = self._pdf_cropper.crop(temp_pdf, pdf_output_path, margins)
            if not crop_result.success:
                return crop_result
            
            # 步骤 2: 调整 DOCX 边距（使用 WPS/Word COM）
            print(f"[DOCX裁剪] 步骤2: 调整 DOCX 边距")
            pages = self._adjust_docx_margins(input_path, output_path, margins)
            
            return CropResult.success_result(
                input_path=input_path,
                output_path=output_path,
                pages=pages,
                orig_res=72,
                out_res=72,  # 边距调整不影响分辨率
            )
            
        except Exception as e:
            logger.error(f"DOCX 裁剪失败: {e}")
            return CropResult.failure_result(input_path, str(e))
        finally:
            # 清理临时文件
            if temp_pdf and os.path.exists(temp_pdf):
                try:
                    os.remove(temp_pdf)
                except Exception:
                    pass
            if temp_dir and os.path.exists(temp_dir):
                try:
                    os.rmdir(temp_dir)
                except Exception:
                    pass

    def _adjust_docx_margins(self, input_path: str, output_path: str, 
                              margins: MarginSettings) -> int:
        """使用 WPS/Word COM 调整 DOCX 边距
        
        Args:
            input_path: 输入 DOCX 路径
            output_path: 输出 DOCX 路径
            margins: 边距调整量（毫米）
            
        Returns:
            页数
        """
        import pythoncom
        import win32com.client
        
        word = None
        doc = None
        
        try:
            # 初始化 COM
            pythoncom.CoInitialize()
            
            # 创建 Word/WPS 应用
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            
            # 转换为绝对路径
            input_path = os.path.abspath(input_path)
            output_path = os.path.abspath(output_path)
            
            # 打开文档
            doc = word.Documents.Open(input_path, ReadOnly=False)
            
            # 获取页数
            page_count = doc.ComputeStatistics(2)  # 2 = wdStatisticPages
            
            # 调整每个节的边距
            # 边距值需要转换为磅（points），1mm = 2.83465 points
            mm_to_points = 2.83465
            
            for section in doc.Sections:
                page_setup = section.PageSetup
                
                # 获取当前边距
                current_top = page_setup.TopMargin
                current_bottom = page_setup.BottomMargin
                current_left = page_setup.LeftMargin
                current_right = page_setup.RightMargin
                
                # 减少边距（裁剪量）
                # 注意：边距不能为负数，最小为 0
                new_top = max(0, current_top - margins.top * mm_to_points)
                new_bottom = max(0, current_bottom - margins.bottom * mm_to_points)
                new_left = max(0, current_left - margins.left * mm_to_points)
                new_right = max(0, current_right - margins.right * mm_to_points)
                
                page_setup.TopMargin = new_top
                page_setup.BottomMargin = new_bottom
                page_setup.LeftMargin = new_left
                page_setup.RightMargin = new_right
                
                logger.info(f"边距调整: 上 {current_top:.1f} → {new_top:.1f}, "
                           f"下 {current_bottom:.1f} → {new_bottom:.1f}, "
                           f"左 {current_left:.1f} → {new_left:.1f}, "
                           f"右 {current_right:.1f} → {new_right:.1f}")
            
            # 另存为新文件
            doc.SaveAs2(output_path)
            
            print(f"[DOCX] 边距调整完成: {output_path}")
            return page_count
            
        except Exception as e:
            logger.error(f"调整 DOCX 边距失败: {e}")
            raise DOCXCropError(input_path, f"调整边距失败: {e}")
        finally:
            # 关闭文档和应用
            try:
                if doc:
                    doc.Close(False)
            except Exception:
                pass
            try:
                if word:
                    word.Quit()
            except Exception:
                pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            time.sleep(0.3)
    
    def get_section_count(self, input_path: str) -> int:
        """获取 DOCX 节数
        
        Args:
            input_path: DOCX 文件路径
            
        Returns:
            节数
        """
        try:
            from docx import Document
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
            from docx import Document
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
