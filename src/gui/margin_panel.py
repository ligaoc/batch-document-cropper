"""边距设置面板"""
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QDoubleSpinBox, QGroupBox
)
from PyQt5.QtCore import pyqtSignal

from ..models.margin_settings import MarginSettings


class MarginPanel(QWidget):
    """边距设置面板"""
    
    # 信号
    margins_changed = pyqtSignal(MarginSettings)
    
    def __init__(self):
        super().__init__()
        self._init_ui()
        self._connect_signals()
    
    def _init_ui(self):
        """初始化 UI"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # 分组框
        group = QGroupBox("边距设置 (毫米)")
        group_layout = QGridLayout(group)
        
        # 上边距
        group_layout.addWidget(QLabel("上:"), 0, 0)
        self._top_spin = QDoubleSpinBox()
        self._top_spin.setRange(0, 100)
        self._top_spin.setValue(0)
        self._top_spin.setSuffix(" mm")
        self._top_spin.setDecimals(1)
        group_layout.addWidget(self._top_spin, 0, 1)
        
        # 下边距
        group_layout.addWidget(QLabel("下:"), 0, 2)
        self._bottom_spin = QDoubleSpinBox()
        self._bottom_spin.setRange(0, 100)
        self._bottom_spin.setValue(0)
        self._bottom_spin.setSuffix(" mm")
        self._bottom_spin.setDecimals(1)
        group_layout.addWidget(self._bottom_spin, 0, 3)
        
        # 左边距
        group_layout.addWidget(QLabel("左:"), 1, 0)
        self._left_spin = QDoubleSpinBox()
        self._left_spin.setRange(0, 100)
        self._left_spin.setValue(0)
        self._left_spin.setSuffix(" mm")
        self._left_spin.setDecimals(1)
        group_layout.addWidget(self._left_spin, 1, 1)
        
        # 右边距
        group_layout.addWidget(QLabel("右:"), 1, 2)
        self._right_spin = QDoubleSpinBox()
        self._right_spin.setRange(0, 100)
        self._right_spin.setValue(0)
        self._right_spin.setSuffix(" mm")
        self._right_spin.setDecimals(1)
        group_layout.addWidget(self._right_spin, 1, 3)
        
        layout.addWidget(group)
    
    def _connect_signals(self):
        """连接信号"""
        self._top_spin.valueChanged.connect(self._on_value_changed)
        self._bottom_spin.valueChanged.connect(self._on_value_changed)
        self._left_spin.valueChanged.connect(self._on_value_changed)
        self._right_spin.valueChanged.connect(self._on_value_changed)
    
    def _on_value_changed(self):
        """值变化"""
        self.margins_changed.emit(self.get_margins())
    
    def get_margins(self) -> MarginSettings:
        """获取边距设置"""
        return MarginSettings(
            top=self._top_spin.value(),
            bottom=self._bottom_spin.value(),
            left=self._left_spin.value(),
            right=self._right_spin.value(),
        )
    
    def set_margins(self, margins: MarginSettings):
        """设置边距值"""
        self._top_spin.blockSignals(True)
        self._bottom_spin.blockSignals(True)
        self._left_spin.blockSignals(True)
        self._right_spin.blockSignals(True)
        
        self._top_spin.setValue(margins.top)
        self._bottom_spin.setValue(margins.bottom)
        self._left_spin.setValue(margins.left)
        self._right_spin.setValue(margins.right)
        
        self._top_spin.blockSignals(False)
        self._bottom_spin.blockSignals(False)
        self._left_spin.blockSignals(False)
        self._right_spin.blockSignals(False)
