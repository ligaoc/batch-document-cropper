"""进度显示组件"""
import os
from typing import List, Dict

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QProgressBar, QGroupBox, QScrollArea
)
from PyQt5.QtCore import Qt


class FileProgressItem(QWidget):
    """单个文件进度项"""
    
    def __init__(self, file_name: str):
        super().__init__()
        self._file_name = file_name
        self._init_ui()
    
    def _init_ui(self):
        """初始化 UI"""
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 2, 0, 2)
        
        # 文件名
        self._name_label = QLabel(self._file_name)
        self._name_label.setMinimumWidth(150)
        self._name_label.setToolTip(self._file_name)
        layout.addWidget(self._name_label)
        
        # 进度条
        self._progress = QProgressBar()
        self._progress.setRange(0, 100)
        self._progress.setValue(0)
        self._progress.setMinimumWidth(100)
        layout.addWidget(self._progress)
        
        # 状态
        self._status_label = QLabel("等待中")
        self._status_label.setMinimumWidth(80)
        layout.addWidget(self._status_label)
    
    def set_progress(self, value: int):
        """设置进度"""
        self._progress.setValue(value)
        if value > 0 and value < 100:
            self._status_label.setText("处理中")
    
    def set_status(self, success: bool, message: str):
        """设置状态"""
        if success:
            self._progress.setValue(100)
            self._status_label.setText("完成")
            self._status_label.setStyleSheet("color: green;")
        else:
            self._status_label.setText("失败")
            self._status_label.setStyleSheet("color: red;")
            self._status_label.setToolTip(message)


class ProgressWidget(QWidget):
    """进度显示组件"""
    
    def __init__(self):
        super().__init__()
        self._items: Dict[str, FileProgressItem] = {}
        self._init_ui()
    
    def _init_ui(self):
        """初始化 UI"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # 分组框
        group = QGroupBox("处理进度")
        group_layout = QVBoxLayout(group)
        
        # 滚动区域
        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setMinimumHeight(100)
        self._scroll.setMaximumHeight(200)
        
        # 内容容器
        self._content = QWidget()
        self._content_layout = QVBoxLayout(self._content)
        self._content_layout.setAlignment(Qt.AlignTop)
        
        self._scroll.setWidget(self._content)
        group_layout.addWidget(self._scroll)
        
        layout.addWidget(group)
    
    def reset(self):
        """重置"""
        # 清空所有项
        for item in self._items.values():
            self._content_layout.removeWidget(item)
            item.deleteLater()
        self._items.clear()
    
    def set_files(self, files: List[str]):
        """设置文件列表"""
        self.reset()
        
        for file_path in files:
            file_name = os.path.basename(file_path)
            item = FileProgressItem(file_name)
            self._items[file_name] = item
            self._content_layout.addWidget(item)
    
    def update_progress(self, file_name: str, progress: int):
        """更新进度"""
        if file_name in self._items:
            self._items[file_name].set_progress(progress)
    
    def set_file_status(self, file_name: str, success: bool, message: str):
        """设置文件状态"""
        if file_name in self._items:
            self._items[file_name].set_status(success, message)
