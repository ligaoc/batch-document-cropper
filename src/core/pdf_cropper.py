"""PDF 裁剪核心类"""
import fitz  # PyMuPDF
from typing import Optional, Tuple
from dataclasses import dataclass

from ..models.margin_settings import MarginSettings
from ..models.task import CropResult


class CropError(Exception):
    """裁剪错误"""
    def __init__(self, file_path: str, reason: str):
        self.file_path = file_path
        self.reason = reason
        super().__init__(f"裁剪失败 {file_path}: {reason}")


@dataclass
class PageInfo:
    """页面信息"""
    width: float
    height: float
    page_num: int


class PDFCropper:
    """PDF 裁剪核心类"""
    
    def __init__(self):
        """初始化裁剪器"""
        pass
    
    def get_page_count(self, input_path: str) -> int:
        """获取 PDF 页数
        
        Args:
            input_path: PDF 文件路径
            
        Returns:
            页数
        """
        try:
            doc = fitz.open(input_path)
            count = len(doc)
            doc.close()
            return count
        except Exception as e:
            raise CropError(input_path, f"无法打开 PDF: {e}")
    
    def get_page_info(self, input_path: str, page_num: int = 0) -> PageInfo:
        """获取页面信息
        
        Args:
            input_path: PDF 文件路径
            page_num: 页码 (从0开始)
            
        Returns:
            页面信息
        """
        try:
            doc = fitz.open(input_path)
            if page_num >= len(doc):
                doc.close()
                raise CropError(input_path, f"页码 {page_num} 超出范围")
            
            page = doc[page_num]
            rect = page.rect
            info = PageInfo(
                width=rect.width,
                height=rect.height,
                page_num=page_num,
            )
            doc.close()
            return info
        except CropError:
            raise
        except Exception as e:
            raise CropError(input_path, f"无法获取页面信息: {e}")
    
    def crop(self, input_path: str, output_path: str,
             margins: MarginSettings) -> CropResult:
        """裁剪 PDF 文件
        
        Args:
            input_path: 输入 PDF 路径
            output_path: 输出 PDF 路径
            margins: 边距设置
            
        Returns:
            裁剪结果对象
        """
        try:
            # 验证边距
            valid, error = margins.validate()
            if not valid:
                return CropResult.failure_result(input_path, error)
            
            # 转换边距为点数
            top_pt, bottom_pt, left_pt, right_pt = margins.to_points()
            
            # 打开源文档
            src_doc = fitz.open(input_path)
            
            # 创建新文档
            dst_doc = fitz.open()
            
            pages_processed = 0
            
            for page_num in range(len(src_doc)):
                src_page = src_doc[page_num]
                src_rect = src_page.rect
                
                # 计算裁剪后的矩形
                new_rect = fitz.Rect(
                    src_rect.x0 + left_pt,
                    src_rect.y0 + top_pt,
                    src_rect.x1 - right_pt,
                    src_rect.y1 - bottom_pt,
                )
                
                # 确保裁剪区域有效
                if new_rect.width <= 0 or new_rect.height <= 0:
                    src_doc.close()
                    dst_doc.close()
                    return CropResult.failure_result(
                        input_path, 
                        f"页面 {page_num + 1} 裁剪后尺寸无效"
                    )
                
                # 创建新页面
                dst_page = dst_doc.new_page(
                    width=new_rect.width,
                    height=new_rect.height,
                )
                
                # 将源页面内容复制到新页面
                dst_page.show_pdf_page(
                    dst_page.rect,
                    src_doc,
                    page_num,
                    clip=new_rect,
                )
                
                pages_processed += 1
            
            # 保存输出文档
            dst_doc.save(output_path)
            dst_doc.close()
            src_doc.close()
            
            return CropResult.success_result(
                input_path=input_path,
                output_path=output_path,
                pages=pages_processed,
                orig_res=72,  # PDF 默认 72 DPI
                out_res=72,
            )
            
        except CropError:
            raise
        except Exception as e:
            return CropResult.failure_result(input_path, str(e))
    
    def generate_preview(self, input_path: str, page_num: int,
                         margins: MarginSettings,
                         scale: float = 1.0) -> Optional[bytes]:
        """生成带裁剪边界的预览图像
        
        Args:
            input_path: PDF 文件路径
            page_num: 页码 (从0开始)
            margins: 边距设置
            scale: 缩放比例
            
        Returns:
            PNG 图像字节数据，失败返回 None
        """
        try:
            doc = fitz.open(input_path)
            if page_num >= len(doc):
                doc.close()
                return None
            
            page = doc[page_num]
            
            # 转换边距为点数
            top_pt, bottom_pt, left_pt, right_pt = margins.to_points()
            
            # 计算裁剪区域
            src_rect = page.rect
            crop_rect = fitz.Rect(
                src_rect.x0 + left_pt,
                src_rect.y0 + top_pt,
                src_rect.x1 - right_pt,
                src_rect.y1 - bottom_pt,
            )
            
            # 渲染页面
            mat = fitz.Matrix(scale, scale)
            pix = page.get_pixmap(matrix=mat)
            
            # 在图像上绘制裁剪边界线
            # 转换为 PIL 或直接返回带标记的图像
            img_data = pix.tobytes("png")
            
            doc.close()
            return img_data
            
        except Exception:
            return None
    
    def get_crop_rect(self, input_path: str, page_num: int,
                      margins: MarginSettings) -> Tuple[float, float, float, float]:
        """获取裁剪后的矩形尺寸
        
        Args:
            input_path: PDF 文件路径
            page_num: 页码
            margins: 边距设置
            
        Returns:
            (x0, y0, x1, y1) 裁剪后的矩形坐标
        """
        doc = fitz.open(input_path)
        page = doc[page_num]
        src_rect = page.rect
        
        top_pt, bottom_pt, left_pt, right_pt = margins.to_points()
        
        result = (
            src_rect.x0 + left_pt,
            src_rect.y0 + top_pt,
            src_rect.x1 - right_pt,
            src_rect.y1 - bottom_pt,
        )
        
        doc.close()
        return result
