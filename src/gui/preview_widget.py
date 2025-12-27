"""预览组件"""
import os
import tempfile
from typing import Optional

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QScrollArea, QGroupBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap, QImage, QPainter, QPen, QColor

from ..models.margin_settings import MarginSettings
from ..core.pdf_cropper import PDFCropper
from ..core.document_converter import DocumentConverter
from ..core.file_validator import needs_conversion


class PreviewWidget(QWidget):
    """预览组件"""
    
    # PDF 默认 72 DPI，1英寸 = 25.4mm
    PDF_DPI = 72.0
    MM_PER_INCH = 25.4
    
    def __init__(self):
        super().__init__()
        self._current_file: Optional[str] = None
        self._current_pdf: Optional[str] = None
        self._temp_pdf: Optional[str] = None
        self._margins = MarginSettings()
        self._current_page = 0
        self._total_pages = 0
        self._cropper = PDFCropper()
        self._converter = DocumentConverter()
        self._scale = 1.0
        self._page_width_pt = 0.0  # 页面宽度（点）
        self._page_height_pt = 0.0  # 页面高度（点）
        
        self._init_ui()
        self._connect_signals()
    
    def _init_ui(self):
        """初始化 UI"""
        layout = QVBoxLayout(self)
        
        # 分组框
        group = QGroupBox("文档预览")
        group_layout = QVBoxLayout(group)
        
        # 预览区域
        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setAlignment(Qt.AlignCenter)
        
        self._preview_label = QLabel("请选择文件进行预览")
        self._preview_label.setAlignment(Qt.AlignCenter)
        self._preview_label.setMinimumSize(400, 500)
        self._preview_label.setStyleSheet("background-color: #f0f0f0;")
        
        self._scroll.setWidget(self._preview_label)
        group_layout.addWidget(self._scroll)
        
        # 页面尺寸信息
        self._size_label = QLabel("页面尺寸: -- × -- mm")
        self._size_label.setAlignment(Qt.AlignCenter)
        self._size_label.setStyleSheet("color: #666; font-size: 11px;")
        group_layout.addWidget(self._size_label)
        
        # 页面导航
        nav_layout = QHBoxLayout()
        self._prev_btn = QPushButton("上一页")
        self._page_label = QLabel("0 / 0")
        self._next_btn = QPushButton("下一页")
        
        nav_layout.addWidget(self._prev_btn)
        nav_layout.addStretch()
        nav_layout.addWidget(self._page_label)
        nav_layout.addStretch()
        nav_layout.addWidget(self._next_btn)
        
        group_layout.addLayout(nav_layout)
        
        layout.addWidget(group)
    
    def _connect_signals(self):
        """连接信号"""
        self._prev_btn.clicked.connect(self._on_prev_page)
        self._next_btn.clicked.connect(self._on_next_page)
    
    def load_document(self, file_path: str, margins: MarginSettings):
        """加载文档
        
        Args:
            file_path: 文件路径
            margins: 边距设置
        """
        self._current_file = file_path
        self._margins = margins
        self._current_page = 0
        
        # 清理之前的临时文件
        self._cleanup_temp()
        
        try:
            # 如果需要转换
            if needs_conversion(file_path):
                if not self._converter.is_available:
                    self._preview_label.setText("LibreOffice 未安装，无法预览 Word 文档")
                    return
                
                temp_dir = tempfile.mkdtemp()
                self._temp_pdf = self._converter.convert_to_pdf(file_path, temp_dir)
                self._current_pdf = self._temp_pdf
            else:
                self._current_pdf = file_path
            
            # 获取页数和页面尺寸
            self._total_pages = self._cropper.get_page_count(self._current_pdf)
            page_info = self._cropper.get_page_info(self._current_pdf, 0)
            self._page_width_pt = page_info.width
            self._page_height_pt = page_info.height
            
            self._update_page_label()
            self._update_size_label()
            self._render_preview()
            
        except Exception as e:
            self._preview_label.setText(f"加载失败: {e}")
    
    def update_margins(self, margins: MarginSettings):
        """更新边距设置"""
        self._margins = margins
        if self._current_pdf:
            self._render_preview()
    
    def _render_preview(self):
        """渲染预览"""
        if not self._current_pdf:
            return
        
        try:
            # 获取当前页面尺寸
            page_info = self._cropper.get_page_info(self._current_pdf, self._current_page)
            self._page_width_pt = page_info.width
            self._page_height_pt = page_info.height
            self._update_size_label()
            
            # 计算适合预览区域的缩放比例
            scroll_width = self._scroll.viewport().width() - 20
            scroll_height = self._scroll.viewport().height() - 20
            
            if scroll_width > 0 and scroll_height > 0:
                scale_x = scroll_width / self._page_width_pt
                scale_y = scroll_height / self._page_height_pt
                self._scale = min(scale_x, scale_y, 2.0)  # 最大2倍
            else:
                self._scale = 1.0
            
            # 生成预览图像
            img_data = self._cropper.generate_preview(
                self._current_pdf,
                self._current_page,
                MarginSettings(),  # 先渲染原始页面
                scale=self._scale
            )
            
            if not img_data:
                self._preview_label.setText("无法生成预览")
                return
            
            # 转换为 QPixmap
            qimg = QImage.fromData(img_data)
            pixmap = QPixmap.fromImage(qimg)
            
            # 绘制裁剪边界线
            pixmap = self._draw_crop_lines(pixmap)
            
            self._preview_label.setPixmap(pixmap)
            
        except Exception as e:
            self._preview_label.setText(f"预览失败: {e}")
    
    def _draw_crop_lines(self, pixmap: QPixmap) -> QPixmap:
        """在预览上绘制裁剪边界线"""
        if self._margins.top == 0 and self._margins.bottom == 0 and \
           self._margins.left == 0 and self._margins.right == 0:
            return pixmap
        
        result = QPixmap(pixmap)
        painter = QPainter(result)
        
        # 设置红色虚线
        pen = QPen(QColor(255, 0, 0))
        pen.setWidth(2)
        pen.setStyle(Qt.DashLine)
        painter.setPen(pen)
        
        # 将边距（毫米）转换为像素
        # 边距(mm) -> 点(pt) -> 像素(px)
        # mm * (72 pt/inch) / (25.4 mm/inch) * scale = px
        mm_to_px = self.PDF_DPI / self.MM_PER_INCH * self._scale
        
        top_px = int(self._margins.top * mm_to_px)
        bottom_px = int(self._margins.bottom * mm_to_px)
        left_px = int(self._margins.left * mm_to_px)
        right_px = int(self._margins.right * mm_to_px)
        
        w = pixmap.width()
        h = pixmap.height()
        
        # 绘制裁剪边界
        if top_px > 0:
            painter.drawLine(0, top_px, w, top_px)
        if bottom_px > 0:
            painter.drawLine(0, h - bottom_px, w, h - bottom_px)
        if left_px > 0:
            painter.drawLine(left_px, 0, left_px, h)
        if right_px > 0:
            painter.drawLine(w - right_px, 0, w - right_px, h)
        
        painter.end()
        return result
    
    def _on_prev_page(self):
        """上一页"""
        if self._current_page > 0:
            self._current_page -= 1
            self._update_page_label()
            self._render_preview()
    
    def _on_next_page(self):
        """下一页"""
        if self._current_page < self._total_pages - 1:
            self._current_page += 1
            self._update_page_label()
            self._render_preview()
    
    def _update_page_label(self):
        """更新页码标签"""
        self._page_label.setText(f"{self._current_page + 1} / {self._total_pages}")
        self._prev_btn.setEnabled(self._current_page > 0)
        self._next_btn.setEnabled(self._current_page < self._total_pages - 1)
    
    def _update_size_label(self):
        """更新页面尺寸标签"""
        # 将点转换为毫米
        width_mm = self._page_width_pt * self.MM_PER_INCH / self.PDF_DPI
        height_mm = self._page_height_pt * self.MM_PER_INCH / self.PDF_DPI
        self._size_label.setText(f"页面尺寸: {width_mm:.1f} × {height_mm:.1f} mm")
    
    def _cleanup_temp(self):
        """清理临时文件"""
        if self._temp_pdf and os.path.exists(self._temp_pdf):
            try:
                os.remove(self._temp_pdf)
                os.rmdir(os.path.dirname(self._temp_pdf))
            except Exception:
                pass
        self._temp_pdf = None
    
    def closeEvent(self, event):
        """关闭事件"""
        self._cleanup_temp()
        super().closeEvent(event)
