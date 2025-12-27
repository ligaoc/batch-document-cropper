"""预览组件"""
import os
import tempfile
from enum import Enum
from typing import Optional

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QScrollArea, QGroupBox, QStackedWidget
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QPixmap, QImage, QPainter, QPen, QColor, QFont

from ..models.margin_settings import MarginSettings
from ..core.pdf_cropper import PDFCropper
from ..core.document_converter import get_document_converter
from ..core.file_validator import needs_conversion


class PreviewState(Enum):
    """预览状态枚举"""
    IDLE = "idle"               # 初始状态，无文件选中
    FILE_SELECTED = "file_selected"  # 文件已选中，等待预览
    LOADING = "loading"         # 加载中
    LOADED = "loaded"           # 已加载
    ERROR = "error"             # 错误


class LoadingIndicator(QWidget):
    """加载指示器组件 - 显示转圈动画"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self._angle = 0
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._rotate)
        self._is_running = False
        
        self.setMinimumSize(100, 100)
        self.setStyleSheet("background-color: transparent;")
    
    def start(self):
        """开始显示加载动画"""
        self._is_running = True
        self._timer.start(50)  # 50ms 更新一次
        self.show()
    
    def stop(self):
        """停止加载动画"""
        self._is_running = False
        self._timer.stop()
        self.hide()
    
    def _rotate(self):
        """旋转动画"""
        self._angle = (self._angle + 10) % 360
        self.update()
    
    def paintEvent(self, event):
        """绘制加载动画"""
        if not self._is_running:
            return
        
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # 计算中心点
        center_x = self.width() // 2
        center_y = self.height() // 2
        radius = min(center_x, center_y) - 10
        
        # 绘制旋转的圆弧
        pen = QPen(QColor(0, 120, 215))  # 蓝色
        pen.setWidth(4)
        pen.setCapStyle(Qt.RoundCap)
        painter.setPen(pen)
        
        # 绘制多个点形成旋转效果
        from math import cos, sin, radians
        for i in range(12):
            angle = radians(self._angle + i * 30)
            alpha = 255 - i * 20
            if alpha < 50:
                alpha = 50
            
            pen.setColor(QColor(0, 120, 215, alpha))
            painter.setPen(pen)
            
            x = center_x + int(radius * cos(angle))
            y = center_y + int(radius * sin(angle))
            painter.drawEllipse(x - 4, y - 4, 8, 8)
        
        # 绘制"加载中"文字
        painter.setPen(QColor(100, 100, 100))
        font = QFont()
        font.setPointSize(10)
        painter.setFont(font)
        painter.drawText(self.rect(), Qt.AlignHCenter | Qt.AlignBottom, "加载中...")
        
        painter.end()


class PreviewLoadingThread(QThread):
    """预览加载线程"""
    
    # 信号: (图像数据, 页面宽度pt, 页面高度pt, 总页数, 临时PDF路径)
    loading_finished = pyqtSignal(bytes, float, float, int, str)
    loading_failed = pyqtSignal(str)  # 错误信息
    
    def __init__(self, file_path: str, page_num: int = 0, scale: float = 1.0):
        super().__init__()
        self._file_path = file_path
        self._page_num = page_num
        self._scale = scale
        self._cropper = PDFCropper()
        self._converter = get_document_converter()
        self._temp_pdf: Optional[str] = None

    def run(self):
        """执行加载"""
        try:
            pdf_path = self._file_path
            temp_pdf = ""
            
            # 如果需要转换 (Word 文档)
            if needs_conversion(self._file_path):
                if not self._converter.is_available:
                    self.loading_failed.emit("没有可用的转换器（需要 Microsoft Word 或 LibreOffice）")
                    return
                
                temp_dir = tempfile.mkdtemp()
                temp_pdf = self._converter.convert_to_pdf(self._file_path, temp_dir)
                pdf_path = temp_pdf
            
            # 获取页数和页面尺寸
            total_pages = self._cropper.get_page_count(pdf_path)
            page_info = self._cropper.get_page_info(pdf_path, self._page_num)
            
            # 生成预览图像
            img_data = self._cropper.generate_preview(
                pdf_path,
                self._page_num,
                MarginSettings(),  # 渲染原始页面
                scale=self._scale
            )
            
            if not img_data:
                self.loading_failed.emit("无法生成预览图像")
                return
            
            self.loading_finished.emit(
                img_data, 
                page_info.width, 
                page_info.height, 
                total_pages,
                temp_pdf
            )
            
        except Exception as e:
            self.loading_failed.emit(str(e))


class PreviewWidget(QWidget):
    """预览组件"""
    
    # PDF 默认 72 DPI，1英寸 = 25.4mm
    PDF_DPI = 72.0
    MM_PER_INCH = 25.4
    
    def __init__(self):
        super().__init__()
        self._state = PreviewState.IDLE
        self._current_file: Optional[str] = None
        self._current_pdf: Optional[str] = None
        self._temp_pdf: Optional[str] = None
        self._margins = MarginSettings()
        self._current_page = 0
        self._total_pages = 0
        self._cropper = PDFCropper()
        self._converter = get_document_converter()
        self._scale = 1.0
        self._page_width_pt = 0.0
        self._page_height_pt = 0.0
        self._cached_pixmap: Optional[QPixmap] = None  # 缓存的预览图像
        self._loading_thread: Optional[PreviewLoadingThread] = None
        
        self._init_ui()
        self._connect_signals()
    
    def _init_ui(self):
        """初始化 UI"""
        layout = QVBoxLayout(self)
        
        # 分组框
        group = QGroupBox("文档预览")
        group_layout = QVBoxLayout(group)
        
        # 预览区域容器 (使用 QStackedWidget 切换不同状态的显示)
        self._preview_container = QWidget()
        container_layout = QVBoxLayout(self._preview_container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        
        # 滚动区域
        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setAlignment(Qt.AlignCenter)
        
        # 预览内容区域
        self._content_widget = QWidget()
        self._content_layout = QVBoxLayout(self._content_widget)
        self._content_layout.setAlignment(Qt.AlignCenter)
        
        # 文件名标签
        self._filename_label = QLabel("")
        self._filename_label.setAlignment(Qt.AlignCenter)
        self._filename_label.setStyleSheet("font-weight: bold; color: #333; margin: 10px;")
        self._filename_label.hide()
        self._content_layout.addWidget(self._filename_label)
        
        # 预览按钮
        self._preview_btn = QPushButton("点击预览")
        self._preview_btn.setMinimumSize(120, 40)
        self._preview_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                border-radius: 5px;
                font-size: 14px;
                padding: 10px 20px;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:disabled {
                background-color: #ccc;
            }
        """)
        self._preview_btn.hide()
        self._content_layout.addWidget(self._preview_btn, alignment=Qt.AlignCenter)
        
        # 加载指示器
        self._loading_indicator = LoadingIndicator()
        self._loading_indicator.hide()
        self._content_layout.addWidget(self._loading_indicator, alignment=Qt.AlignCenter)
        
        # 预览图像标签
        self._preview_label = QLabel("请选择文件进行预览")
        self._preview_label.setAlignment(Qt.AlignCenter)
        self._preview_label.setMinimumSize(400, 400)
        self._preview_label.setStyleSheet("background-color: #f0f0f0;")
        self._content_layout.addWidget(self._preview_label)
        
        # 错误信息标签
        self._error_label = QLabel("")
        self._error_label.setAlignment(Qt.AlignCenter)
        self._error_label.setStyleSheet("color: red; margin: 10px;")
        self._error_label.hide()
        self._content_layout.addWidget(self._error_label)
        
        self._scroll.setWidget(self._content_widget)
        container_layout.addWidget(self._scroll)
        
        group_layout.addWidget(self._preview_container)
        
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
        
        self._prev_btn.setEnabled(False)
        self._next_btn.setEnabled(False)
        
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
        self._preview_btn.clicked.connect(self._start_preview)

    def on_file_selected(self, file_path: str, margins: MarginSettings):
        """文件选中时调用 - 不自动加载预览
        
        Args:
            file_path: 选中的文件路径
            margins: 当前边距设置
        """
        # 清理之前的状态
        self._cleanup_temp()
        self._cached_pixmap = None
        self._current_pdf = None
        
        self._current_file = file_path
        self._margins = margins
        self._current_page = 0
        self._total_pages = 0
        
        # 更新状态
        self._state = PreviewState.FILE_SELECTED
        
        # 显示文件名和预览按钮
        filename = os.path.basename(file_path)
        self._filename_label.setText(f"已选择: {filename}")
        self._filename_label.show()
        
        self._preview_btn.setEnabled(True)
        self._preview_btn.show()
        
        # 隐藏其他元素
        self._loading_indicator.stop()
        self._error_label.hide()
        self._preview_label.setText("点击上方按钮加载预览")
        self._preview_label.setPixmap(QPixmap())  # 清除之前的图像
        
        # 重置页面信息
        self._page_label.setText("0 / 0")
        self._size_label.setText("页面尺寸: -- × -- mm")
        self._prev_btn.setEnabled(False)
        self._next_btn.setEnabled(False)
    
    def _start_preview(self):
        """开始加载预览"""
        if not self._current_file:
            return
        
        # 更新状态
        self._state = PreviewState.LOADING
        
        # 显示加载动画，禁用按钮
        self._preview_btn.setEnabled(False)
        self._preview_btn.hide()
        self._loading_indicator.start()
        self._error_label.hide()
        self._preview_label.setText("")
        
        # 计算缩放比例
        scroll_width = self._scroll.viewport().width() - 20
        scroll_height = self._scroll.viewport().height() - 20
        # 使用默认缩放，加载完成后会重新计算
        scale = 1.0
        
        # 启动加载线程
        self._loading_thread = PreviewLoadingThread(
            self._current_file, 
            self._current_page, 
            scale
        )
        self._loading_thread.loading_finished.connect(self._on_loading_finished)
        self._loading_thread.loading_failed.connect(self._on_loading_failed)
        self._loading_thread.start()
    
    def _on_loading_finished(self, img_data: bytes, width_pt: float, height_pt: float, 
                              total_pages: int, temp_pdf: str):
        """加载完成处理"""
        self._state = PreviewState.LOADED
        
        # 保存临时 PDF 路径
        if temp_pdf:
            self._temp_pdf = temp_pdf
            self._current_pdf = temp_pdf
        else:
            self._current_pdf = self._current_file
        
        # 保存页面信息
        self._page_width_pt = width_pt
        self._page_height_pt = height_pt
        self._total_pages = total_pages
        
        # 停止加载动画
        self._loading_indicator.stop()
        self._preview_btn.hide()
        self._filename_label.hide()
        
        # 更新页面信息
        self._update_page_label()
        self._update_size_label()
        
        # 使用正确的缩放比例重新渲染预览
        self._render_preview()
    
    def _on_loading_failed(self, error_message: str):
        """加载失败处理"""
        self._state = PreviewState.ERROR
        
        # 停止加载动画
        self._loading_indicator.stop()
        
        # 显示错误信息和重试按钮
        self._error_label.setText(f"预览失败: {error_message}")
        self._error_label.show()
        
        self._preview_btn.setText("重试")
        self._preview_btn.setEnabled(True)
        self._preview_btn.show()
        
        self._preview_label.setText("")
    
    def update_margins(self, margins: MarginSettings):
        """更新边距设置 - 仅重绘裁剪线，不重新加载文档"""
        self._margins = margins
        
        # 只有在已加载状态下才重绘
        if self._state == PreviewState.LOADED and self._cached_pixmap:
            pixmap = self._draw_crop_lines(self._cached_pixmap)
            self._preview_label.setPixmap(pixmap)
    
    # 保留旧方法名以兼容
    def load_document(self, file_path: str, margins: MarginSettings):
        """加载文档 - 兼容旧接口，现在改为只选中文件"""
        self.on_file_selected(file_path, margins)

    def _render_preview(self):
        """渲染预览 - 用于翻页时重新加载"""
        if not self._current_pdf or self._state != PreviewState.LOADED:
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
                self._scale = min(scale_x, scale_y, 2.0)
            else:
                self._scale = 1.0
            
            # 生成预览图像
            img_data = self._cropper.generate_preview(
                self._current_pdf,
                self._current_page,
                MarginSettings(),
                scale=self._scale
            )
            
            if not img_data:
                self._preview_label.setText("无法生成预览")
                return
            
            # 转换为 QPixmap
            qimg = QImage.fromData(img_data)
            self._cached_pixmap = QPixmap.fromImage(qimg)
            
            # 绘制裁剪边界线
            pixmap = self._draw_crop_lines(self._cached_pixmap)
            self._preview_label.setPixmap(pixmap)
            
        except Exception as e:
            self._preview_label.setText(f"预览失败: {e}")
    
    def _draw_crop_lines(self, pixmap: QPixmap) -> QPixmap:
        """在预览上绘制裁剪边界线"""
        if not pixmap or pixmap.isNull():
            return pixmap
            
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
