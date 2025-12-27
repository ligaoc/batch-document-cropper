"""批量处理器"""
import os
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Optional
from dataclasses import dataclass

from PyQt5.QtCore import QObject, pyqtSignal

from ..models.margin_settings import MarginSettings
from ..models.task import ProcessingTask, TaskStatus, CropResult, ProcessingSummary
from .document_converter import get_document_converter, ConversionError
from .pdf_cropper import PDFCropper
from .docx_cropper import DOCXCropper
from .file_validator import is_pdf, is_docx, is_doc, get_output_extension


@dataclass
class BatchTask:
    """批处理任务"""
    file_path: str
    margins: MarginSettings
    output_dir: str
    output_suffix: str = "_cropped"


class BatchProcessor(QObject):
    """批量处理器，管理多文档并发处理
    
    处理逻辑:
    - PDF 文件: 使用 PDFCropper 裁剪，输出 PDF
    - DOCX/DOC 文件: 转 PDF → 裁剪 → 输出 PDF + DOCX（图片版）
    """
    
    # Qt 信号
    progress_updated = pyqtSignal(str, int)  # (文件名, 进度)
    file_completed = pyqtSignal(str, bool, str)  # (文件名, 是否成功, 消息)
    all_completed = pyqtSignal(int, int)  # (成功数, 失败数)
    
    MAX_WORKERS = 5  # 最大并发数
    
    def __init__(self, max_workers: int = 5):
        """初始化批量处理器
        
        Args:
            max_workers: 最大并发数 (最大为5)
        """
        super().__init__()
        self._max_workers = min(max_workers, self.MAX_WORKERS)
        self._tasks: List[BatchTask] = []
        self._converter = get_document_converter()
        self._pdf_cropper = PDFCropper()
        self._docx_cropper = DOCXCropper()
        self._cancelled = False
        self._executor: Optional[ThreadPoolExecutor] = None
    
    @property
    def task_count(self) -> int:
        """当前任务数"""
        return len(self._tasks)
    
    def add_task(self, file_path: str, margins: MarginSettings,
                 output_dir: str, output_suffix: str = "_cropped") -> bool:
        """添加处理任务
        
        Args:
            file_path: 文件路径
            margins: 边距设置
            output_dir: 输出目录
            output_suffix: 输出文件后缀
            
        Returns:
            是否成功添加
        """
        if len(self._tasks) >= self.MAX_WORKERS:
            return False
        
        self._tasks.append(BatchTask(
            file_path=file_path,
            margins=margins,
            output_dir=output_dir,
            output_suffix=output_suffix,
        ))
        return True
    
    def clear_tasks(self) -> None:
        """清空任务队列"""
        self._tasks.clear()
    
    def _process_single_file(self, task: BatchTask) -> CropResult:
        """处理单个文件
        
        Args:
            task: 批处理任务
            
        Returns:
            处理结果
        """
        file_path = task.file_path
        file_name = os.path.basename(file_path)
        
        try:
            # 发送进度: 开始
            self.progress_updated.emit(file_name, 10)
            
            # 确定输出扩展名
            output_ext = get_output_extension(file_path)
            
            # 构建输出路径
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            output_path = os.path.join(
                task.output_dir,
                f"{base_name}{task.output_suffix}{output_ext}"
            )
            
            # 确保输出目录存在
            os.makedirs(task.output_dir, exist_ok=True)
            
            # 根据文件类型选择处理方式
            if is_pdf(file_path):
                # PDF 文件: 直接裁剪
                result = self._process_pdf(file_path, output_path, task.margins, file_name)
            elif is_docx(file_path):
                # DOCX 文件: 直接裁剪
                result = self._process_docx(file_path, output_path, task.margins, file_name)
            elif is_doc(file_path):
                # DOC 文件: 先转换为 DOCX，再裁剪
                result = self._process_doc(file_path, output_path, task.margins, file_name)
            else:
                result = CropResult.failure_result(file_path, "不支持的文件格式")
            
            self.progress_updated.emit(file_name, 100)
            return result
            
        except Exception as e:
            return CropResult.failure_result(file_path, str(e))
    
    def _process_pdf(self, input_path: str, output_path: str,
                     margins: MarginSettings, file_name: str) -> CropResult:
        """处理 PDF 文件
        
        Args:
            input_path: 输入路径
            output_path: 输出路径
            margins: 边距设置
            file_name: 文件名（用于进度更新）
            
        Returns:
            处理结果
        """
        self.progress_updated.emit(file_name, 30)
        result = self._pdf_cropper.crop(input_path, output_path, margins)
        self.progress_updated.emit(file_name, 90)
        return result
    
    def _process_docx(self, input_path: str, output_path: str,
                      margins: MarginSettings, file_name: str) -> CropResult:
        """处理 DOCX 文件
        
        Args:
            input_path: 输入路径
            output_path: 输出路径
            margins: 边距设置
            file_name: 文件名（用于进度更新）
            
        Returns:
            处理结果
        """
        self.progress_updated.emit(file_name, 30)
        result = self._docx_cropper.crop(input_path, output_path, margins)
        self.progress_updated.emit(file_name, 90)
        return result
    
    def _process_doc(self, input_path: str, output_path: str,
                     margins: MarginSettings, file_name: str) -> CropResult:
        """处理 DOC 文件（先转换为 DOCX，再裁剪）
        
        Args:
            input_path: 输入路径
            output_path: 输出路径
            margins: 边距设置
            file_name: 文件名（用于进度更新）
            
        Returns:
            处理结果
        """
        temp_docx = None
        temp_dir = None
        
        try:
            self.progress_updated.emit(file_name, 20)
            
            # 创建临时目录
            temp_dir = tempfile.mkdtemp()
            
            # 转换 DOC 到 DOCX
            temp_docx = self._converter.convert_doc_to_docx(input_path, temp_dir)
            
            self.progress_updated.emit(file_name, 50)
            
            # 裁剪 DOCX
            result = self._docx_cropper.crop(temp_docx, output_path, margins)
            
            self.progress_updated.emit(file_name, 90)
            return result
            
        except ConversionError as e:
            return CropResult.failure_result(input_path, str(e))
        finally:
            # 清理临时文件
            if temp_docx and os.path.exists(temp_docx):
                try:
                    os.remove(temp_docx)
                except Exception:
                    pass
            if temp_dir and os.path.exists(temp_dir):
                try:
                    os.rmdir(temp_dir)
                except Exception:
                    pass
    
    def start(self) -> ProcessingSummary:
        """开始处理所有任务
        
        Returns:
            处理摘要
        """
        self._cancelled = False
        summary = ProcessingSummary()
        start_time = time.time()
        
        if not self._tasks:
            return summary
        
        with ThreadPoolExecutor(max_workers=self._max_workers) as executor:
            self._executor = executor
            
            # 提交所有任务
            future_to_task = {
                executor.submit(self._process_single_file, task): task
                for task in self._tasks
            }
            
            # 收集结果
            for future in as_completed(future_to_task):
                if self._cancelled:
                    break
                
                task = future_to_task[future]
                file_name = os.path.basename(task.file_path)
                
                try:
                    result = future.result()
                    
                    if result.success:
                        summary.add_success()
                        self.file_completed.emit(file_name, True, "处理成功")
                    else:
                        summary.add_failure(task.file_path)
                        self.file_completed.emit(file_name, False, result.error_message)
                        
                except Exception as e:
                    summary.add_failure(task.file_path)
                    self.file_completed.emit(file_name, False, str(e))
        
        self._executor = None
        summary.total_time = time.time() - start_time
        
        # 清空任务队列
        self._tasks.clear()
        
        # 发送完成信号
        self.all_completed.emit(summary.successful, summary.failed)
        
        return summary
    
    def cancel(self) -> None:
        """取消处理"""
        self._cancelled = True
        if self._executor:
            self._executor.shutdown(wait=False, cancel_futures=True)
