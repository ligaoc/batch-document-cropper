"""文件列表组件"""
import os
from typing import List, Optional

from PyQt5.QtWidgets import (
    QListWidget, QListWidgetItem, QAbstractItemView
)
from PyQt5.QtCore import pyqtSignal, Qt

from ..core.file_validator import is_supported_format


class FileListWidget(QListWidget):
    """文件列表组件"""
    
    MAX_FILES = 5
    
    # 信号
    selection_changed = pyqtSignal(str)  # 选中的文件路径
    
    def __init__(self):
        super().__init__()
        self._files: List[str] = []
        
        self._init_ui()
        self._connect_signals()
    
    def _init_ui(self):
        """初始化 UI"""
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.DropOnly)
        self.setSelectionMode(QAbstractItemView.SingleSelection)
        self.setMinimumHeight(150)
    
    def _connect_signals(self):
        """连接信号"""
        self.itemSelectionChanged.connect(self._on_selection_changed)
    
    def _on_selection_changed(self):
        """选择变化"""
        items = self.selectedItems()
        if items:
            index = self.row(items[0])
            if 0 <= index < len(self._files):
                self.selection_changed.emit(self._files[index])
        else:
            self.selection_changed.emit("")
    
    def add_file(self, file_path: str) -> bool:
        """添加文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            是否成功添加
        """
        if len(self._files) >= self.MAX_FILES:
            return False
        
        if not is_supported_format(file_path):
            return False
        
        if file_path in self._files:
            return False
        
        self._files.append(file_path)
        
        # 添加列表项
        item = QListWidgetItem(os.path.basename(file_path))
        item.setToolTip(file_path)
        self.addItem(item)
        
        return True
    
    def remove_selected(self):
        """移除选中的文件"""
        items = self.selectedItems()
        for item in items:
            index = self.row(item)
            if 0 <= index < len(self._files):
                self._files.pop(index)
                self.takeItem(index)
    
    def clear_files(self):
        """清空文件列表"""
        self._files.clear()
        self.clear()
    
    def get_files(self) -> List[str]:
        """获取所有文件路径"""
        return self._files.copy()
    
    def get_selected_file(self) -> Optional[str]:
        """获取选中的文件路径"""
        items = self.selectedItems()
        if items:
            index = self.row(items[0])
            if 0 <= index < len(self._files):
                return self._files[index]
        return None
    
    def dragEnterEvent(self, event):
        """拖入事件"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()
    
    def dragMoveEvent(self, event):
        """拖动事件"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()
    
    def dropEvent(self, event):
        """放下事件"""
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if is_supported_format(file_path):
                    if not self.add_file(file_path):
                        break
            event.acceptProposedAction()
        else:
            event.ignore()
