"""DOCX 裁剪器 - 通过转换为 PDF 裁剪后再转回 DOCX（图片版）"""
import os
import tempfile
import io
import logging
from typing import Optional, Tuple

import fitz  # PyMuPDF
from docx import Document
from docx.shared import Mm, Inches
from docx.enum.section import WD_ORIENT

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
    1. DOCX → PDF（优先使用 Microsoft Word，备选 LibreOffice）
    2. PDF → 裁剪后的 PDF（使用 PDFCropper）
    3. 裁剪后的 PDF → 高清图片 → 新 DOCX（每页一张图片）
    
    输出：
    - _cropped.pdf: 裁剪后的 PDF
    - _cropped.docx: 裁剪后的 DOCX（图片版）
    """
    
    # 图片 DPI，越高越清晰，但文件越大
    IMAGE_DPI = 200
    
    def __init__(self):
        """初始化裁剪器"""
        self._converter = get_document_converter()
        self._pdf_cropper = PDFCropper()
        print(f"[DOCX裁剪器] 初始化完成，转换器: {self._converter.converter_name}")
    
    def crop(self, input_path: str, output_path: str,
             margins: MarginSettings) -> CropResult:
        """裁剪 DOCX 文件
        
        会同时输出 PDF 和 DOCX 两种格式：
        - output_path: DOCX 文件（图片版）
        - output_path 同目录下的 .pdf 文件
        
        Args:
            input_path: 输入 DOCX 路径
            output_path: 输出 DOCX 路径
            margins: 裁剪量设置（从页面边缘裁剪掉的毫米数）
            
        Returns:
            裁剪结果对象
        """
        temp_dir = None
        temp_pdf = None
        cropped_pdf = None
        
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
                    "没有可用的转换器（需要 Microsoft Word 或 LibreOffice）"
                )
            
            # 创建临时目录
            temp_dir = tempfile.mkdtemp()
            
            # 步骤 1: DOCX → PDF
            print(f"[DOCX裁剪] 步骤1: 转换为 PDF，使用 {self._converter.converter_name}")
            logger.info(f"DOCX → PDF 转换，使用: {self._converter.converter_name}")
            temp_pdf = self._converter.convert_to_pdf(input_path, temp_dir)
            
            # 步骤 2: 裁剪 PDF
            # 确定输出路径
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            
            # PDF 输出路径（与 DOCX 同名但扩展名为 .pdf）
            base_name = os.path.splitext(output_path)[0]
            pdf_output_path = base_name + ".pdf"
            
            # 裁剪 PDF
            crop_result = self._pdf_cropper.crop(temp_pdf, pdf_output_path, margins)
            if not crop_result.success:
                return crop_result
            
            # 步骤 3: 裁剪后的 PDF → DOCX（图片版）
            pages = self._pdf_to_image_docx(pdf_output_path, output_path)
            
            return CropResult.success_result(
                input_path=input_path,
                output_path=output_path,
                pages=pages,
                orig_res=72,
                out_res=self.IMAGE_DPI,
            )
            
        except Exception as e:
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

    def _pdf_to_image_docx(self, pdf_path: str, docx_path: str) -> int:
        """将 PDF 转换为图片版 DOCX
        
        每页 PDF 转为一张高清图片，插入到新的 DOCX 文档中。
        
        Args:
            pdf_path: 输入 PDF 路径
            docx_path: 输出 DOCX 路径
            
        Returns:
            页数
        """
        # 打开 PDF
        pdf_doc = fitz.open(pdf_path)
        page_count = len(pdf_doc)
        
        # 创建新的 DOCX 文档
        doc = Document()
        
        # 计算缩放比例（DPI / 72，因为 PDF 默认 72 DPI）
        zoom = self.IMAGE_DPI / 72.0
        mat = fitz.Matrix(zoom, zoom)
        
        for page_num in range(page_count):
            page = pdf_doc[page_num]
            
            # 获取页面尺寸（点数）
            page_rect = page.rect
            page_width_pt = page_rect.width
            page_height_pt = page_rect.height
            
            # 转换为毫米
            page_width_mm = page_width_pt * 25.4 / 72
            page_height_mm = page_height_pt * 25.4 / 72
            
            # 渲染页面为图片
            pix = page.get_pixmap(matrix=mat)
            img_data = pix.tobytes("png")
            
            # 设置页面尺寸和边距
            if page_num == 0:
                section = doc.sections[0]
            else:
                section = doc.add_section()
            
            # 设置页面尺寸
            section.page_width = Mm(page_width_mm)
            section.page_height = Mm(page_height_mm)
            
            # 设置最小边距（让图片占满页面）
            section.top_margin = Mm(0)
            section.bottom_margin = Mm(0)
            section.left_margin = Mm(0)
            section.right_margin = Mm(0)
            
            # 判断页面方向
            if page_width_mm > page_height_mm:
                section.orientation = WD_ORIENT.LANDSCAPE
            else:
                section.orientation = WD_ORIENT.PORTRAIT
            
            # 添加图片
            if page_num > 0:
                # 非第一页需要先添加段落
                pass
            
            # 将图片数据写入内存流
            img_stream = io.BytesIO(img_data)
            
            # 添加图片到文档，尺寸与页面一致
            paragraph = doc.add_paragraph()
            run = paragraph.add_run()
            run.add_picture(img_stream, width=Mm(page_width_mm), height=Mm(page_height_mm))
            
            # 设置段落无间距
            paragraph.paragraph_format.space_before = Mm(0)
            paragraph.paragraph_format.space_after = Mm(0)
            paragraph.paragraph_format.line_spacing = 1.0
        
        pdf_doc.close()
        
        # 保存 DOCX
        doc.save(docx_path)
        
        return page_count
    
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
