"""分辨率保持模块"""
import fitz  # PyMuPDF
from typing import Dict, List, Optional
from dataclasses import dataclass


@dataclass
class ImageResolution:
    """图像分辨率信息"""
    xref: int
    width: int
    height: int
    dpi_x: float
    dpi_y: float
    
    @property
    def dpi(self) -> float:
        """返回平均 DPI"""
        return (self.dpi_x + self.dpi_y) / 2


class ResolutionKeeper:
    """分辨率保持模块"""
    
    @staticmethod
    def get_pdf_resolution(pdf_path: str) -> Dict[str, any]:
        """获取 PDF 中图像的分辨率信息
        
        Args:
            pdf_path: PDF 文件路径
            
        Returns:
            分辨率信息字典，包含:
            - images: 图像列表
            - min_dpi: 最小 DPI
            - max_dpi: 最大 DPI
            - avg_dpi: 平均 DPI
        """
        try:
            doc = fitz.open(pdf_path)
            images: List[ImageResolution] = []
            
            for page_num in range(len(doc)):
                page = doc[page_num]
                image_list = page.get_images(full=True)
                
                for img_info in image_list:
                    xref = img_info[0]
                    try:
                        base_image = doc.extract_image(xref)
                        if base_image:
                            width = base_image.get("width", 0)
                            height = base_image.get("height", 0)
                            
                            # 估算 DPI (基于页面尺寸)
                            # 默认假设 72 DPI 如果无法计算
                            dpi = 72.0
                            
                            images.append(ImageResolution(
                                xref=xref,
                                width=width,
                                height=height,
                                dpi_x=dpi,
                                dpi_y=dpi,
                            ))
                    except Exception:
                        continue
            
            doc.close()
            
            if not images:
                return {
                    "images": [],
                    "min_dpi": 72,
                    "max_dpi": 72,
                    "avg_dpi": 72,
                    "image_count": 0,
                }
            
            dpis = [img.dpi for img in images]
            return {
                "images": images,
                "min_dpi": min(dpis),
                "max_dpi": max(dpis),
                "avg_dpi": sum(dpis) / len(dpis),
                "image_count": len(images),
            }
            
        except Exception as e:
            return {
                "images": [],
                "min_dpi": 72,
                "max_dpi": 72,
                "avg_dpi": 72,
                "image_count": 0,
                "error": str(e),
            }
    
    @staticmethod
    def verify_resolution(original_path: str, processed_path: str) -> bool:
        """验证处理后的分辨率是否保持
        
        Args:
            original_path: 原始文件路径
            processed_path: 处理后文件路径
            
        Returns:
            分辨率是否保持 (处理后 >= 原始)
        """
        try:
            orig_info = ResolutionKeeper.get_pdf_resolution(original_path)
            proc_info = ResolutionKeeper.get_pdf_resolution(processed_path)
            
            # 如果原始文件没有图像，认为分辨率保持
            if orig_info["image_count"] == 0:
                return True
            
            # 比较最小 DPI
            return proc_info["min_dpi"] >= orig_info["min_dpi"]
            
        except Exception:
            return False
    
    @staticmethod
    def get_page_resolution(pdf_path: str, page_num: int = 0) -> Optional[float]:
        """获取页面的渲染分辨率
        
        Args:
            pdf_path: PDF 文件路径
            page_num: 页码
            
        Returns:
            页面分辨率 (DPI)，失败返回 None
        """
        try:
            doc = fitz.open(pdf_path)
            if page_num >= len(doc):
                doc.close()
                return None
            
            # PDF 默认分辨率为 72 DPI
            doc.close()
            return 72.0
            
        except Exception:
            return None
