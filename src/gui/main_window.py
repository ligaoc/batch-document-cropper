"""主窗口"""
import os
from typing import Optional

from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QMessageBox, QSplitter,
    QStatusBar, QProgressBar, QLabel
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

from ..models.margin_settings import MarginSettings
from ..models.task import ProcessingSummary
from ..core.batch_processor import BatchProcessor
from ..core.file_validator import is_supported_format
from .file_list_widget import FileListWidget
from .margin_panel import MarginPanel
from .preview_widget import PreviewWidget
from .progress_widget import ProgressWidget


class ProcessingThread(QThread):
    """处理线程"""
    finished = pyqtSignal(object)  # ProcessingSummary
    
    def __init__(self, processor: BatchProcessor):
        super().__init__()
        self._processor = processor
    
    def run(self):
        summary = self._processor.start()
        self.finished.emit(summary)


class MainWindow(QMainWindow):
    """主窗口类"""
    
    def __init__(self):
        super().__init__()
        self._processor = BatchProcessor()
        self._processing_thread: Optional[ProcessingThread] = None
        self._output_dir = ""
        
        self._init_ui()
        self._connect_signals()
    
    def _init_ui(self):
        """初始化 UI"""
        self.setWindowTitle("批量文档裁剪工具")
        self.setMinimumSize(900, 600)
        
        # 中央部件
        central = QWidget()
        self.setCentralWidget(central)
        
        # 主布局
        main_layout = QHBoxLayout(central)
        
        # 左侧面板
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        
        # 文件列表
        self._file_list = FileListWidget()
        left_layout.addWidget(QLabel("文件列表 (最多5个):"))
        left_layout.addWidget(self._file_list)
        
        # 文件操作按钮
        file_btn_layout = QHBoxLayout()
        self._add_btn = QPushButton("添加文件")
        self._remove_btn = QPushButton("移除选中")
        self._clear_btn = QPushButton("清空列表")
        file_btn_layout.addWidget(self._add_btn)
        file_btn_layout.addWidget(self._remove_btn)
        file_btn_layout.addWidget(self._clear_btn)
        left_layout.addLayout(file_btn_layout)
        
        # 边距设置
        self._margin_panel = MarginPanel()
        left_layout.addWidget(self._margin_panel)
        
        # 输出目录
        output_layout = QHBoxLayout()
        self._output_label = QLabel("输出目录: 未选择")
        self._output_btn = QPushButton("选择目录")
        output_layout.addWidget(self._output_label, 1)
        output_layout.addWidget(self._output_btn)
        left_layout.addLayout(output_layout)
        
        # 处理按钮
        btn_layout = QHBoxLayout()
        self._start_btn = QPushButton("开始处理")
        self._cancel_btn = QPushButton("取消")
        self._cancel_btn.setEnabled(False)
        btn_layout.addWidget(self._start_btn)
        btn_layout.addWidget(self._cancel_btn)
        left_layout.addLayout(btn_layout)
        
        # 进度显示
        self._progress_widget = ProgressWidget()
        left_layout.addWidget(self._progress_widget)
        
        # 右侧预览
        self._preview = PreviewWidget()
        
        # 分割器
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(left_panel)
        splitter.addWidget(self._preview)
        splitter.setSizes([400, 500])
        
        main_layout.addWidget(splitter)
        
        # 状态栏
        self._status_bar = QStatusBar()
        self.setStatusBar(self._status_bar)
        self._status_bar.showMessage("就绪")
    
    def _connect_signals(self):
        """连接信号"""
        self._add_btn.clicked.connect(self._on_add_files)
        self._remove_btn.clicked.connect(self._file_list.remove_selected)
        self._clear_btn.clicked.connect(self._file_list.clear_files)
        self._output_btn.clicked.connect(self._on_select_output)
        self._start_btn.clicked.connect(self._on_start)
        self._cancel_btn.clicked.connect(self._on_cancel)
        
        # 文件选择变化时更新预览
        self._file_list.selection_changed.connect(self._on_file_selected)
        
        # 边距变化时更新预览
        self._margin_panel.margins_changed.connect(self._on_margins_changed)
        
        # 处理器信号
        self._processor.progress_updated.connect(self._progress_widget.update_progress)
        self._processor.file_completed.connect(self._on_file_completed)
        self._processor.all_completed.connect(self._on_all_completed)
    
    def _on_add_files(self):
        """添加文件"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择文档",
            "",
            "支持的文档 (*.pdf *.docx *.doc);;PDF 文件 (*.pdf);;Word 文档 (*.docx *.doc)"
        )
        
        for file_path in files:
            if is_supported_format(file_path):
                if not self._file_list.add_file(file_path):
                    QMessageBox.warning(self, "提示", "最多只能添加5个文件")
                    break
    
    def _on_select_output(self):
        """选择输出目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if dir_path:
            self._output_dir = dir_path
            self._output_label.setText(f"输出目录: {dir_path}")
    
    def _on_file_selected(self, file_path: str):
        """文件选择变化 - 不自动加载预览，只通知预览组件"""
        if file_path:
            margins = self._margin_panel.get_margins()
            self._preview.on_file_selected(file_path, margins)
    
    def _on_margins_changed(self, margins: MarginSettings):
        """边距变化"""
        current_file = self._file_list.get_selected_file()
        if current_file:
            self._preview.update_margins(margins)
    
    def _on_start(self):
        """开始处理"""
        files = self._file_list.get_files()
        if not files:
            QMessageBox.warning(self, "提示", "请先添加文件")
            return
        
        if not self._output_dir:
            QMessageBox.warning(self, "提示", "请先选择输出目录")
            return
        
        margins = self._margin_panel.get_margins()
        valid, error = margins.validate()
        if not valid:
            QMessageBox.warning(self, "边距错误", error)
            return
        
        # 清空处理器任务
        self._processor.clear_tasks()
        
        # 添加任务
        for file_path in files:
            self._processor.add_task(file_path, margins, self._output_dir)
        
        # 更新 UI 状态
        self._start_btn.setEnabled(False)
        self._cancel_btn.setEnabled(True)
        self._progress_widget.reset()
        self._progress_widget.set_files(files)
        self._status_bar.showMessage("处理中...")
        
        # 启动处理线程
        self._processing_thread = ProcessingThread(self._processor)
        self._processing_thread.finished.connect(self._on_processing_finished)
        self._processing_thread.start()
    
    def _on_cancel(self):
        """取消处理"""
        self._processor.cancel()
        self._status_bar.showMessage("正在取消...")
    
    def _on_file_completed(self, file_name: str, success: bool, message: str):
        """单个文件处理完成"""
        self._progress_widget.set_file_status(file_name, success, message)
    
    def _on_all_completed(self, success_count: int, fail_count: int):
        """所有文件处理完成"""
        pass  # 由 _on_processing_finished 处理
    
    def _on_processing_finished(self, summary: ProcessingSummary):
        """处理线程完成"""
        self._start_btn.setEnabled(True)
        self._cancel_btn.setEnabled(False)
        
        msg = f"处理完成: 成功 {summary.successful} 个, 失败 {summary.failed} 个"
        self._status_bar.showMessage(msg)
        
        if summary.failed > 0:
            failed_list = "\n".join(os.path.basename(f) for f in summary.failed_files)
            QMessageBox.warning(
                self, 
                "处理完成", 
                f"{msg}\n\n失败的文件:\n{failed_list}"
            )
        else:
            QMessageBox.information(self, "处理完成", msg)
