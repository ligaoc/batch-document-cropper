"""任务和结果数据模型"""
from dataclasses import dataclass, field
from enum import Enum
from typing import List
import uuid

from .margin_settings import MarginSettings


class TaskStatus(Enum):
    """任务状态枚举"""
    PENDING = "pending"
    CONVERTING = "converting"
    CROPPING = "cropping"
    COMPLETED = "completed"
    FAILED = "failed"


@dataclass
class ProcessingTask:
    """处理任务数据模型"""
    input_path: str
    output_path: str
    margins: MarginSettings
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    status: TaskStatus = TaskStatus.PENDING
    progress: int = 0
    error_message: str = ""
    
    def set_status(self, status: TaskStatus) -> None:
        """设置任务状态"""
        self.status = status
    
    def set_progress(self, progress: int) -> None:
        """设置进度 (0-100)"""
        self.progress = max(0, min(100, progress))
    
    def set_error(self, message: str) -> None:
        """设置错误信息并标记为失败"""
        self.error_message = message
        self.status = TaskStatus.FAILED


@dataclass
class CropResult:
    """裁剪结果数据模型"""
    success: bool
    input_path: str
    output_path: str
    pages_processed: int = 0
    original_resolution: int = 0
    output_resolution: int = 0
    error_message: str = ""
    
    @classmethod
    def success_result(cls, input_path: str, output_path: str,
                       pages: int, orig_res: int, out_res: int) -> "CropResult":
        """创建成功结果"""
        return cls(
            success=True,
            input_path=input_path,
            output_path=output_path,
            pages_processed=pages,
            original_resolution=orig_res,
            output_resolution=out_res,
        )
    
    @classmethod
    def failure_result(cls, input_path: str, error: str) -> "CropResult":
        """创建失败结果"""
        return cls(
            success=False,
            input_path=input_path,
            output_path="",
            error_message=error,
        )


@dataclass
class ProcessingSummary:
    """批量处理摘要"""
    total_files: int = 0
    successful: int = 0
    failed: int = 0
    failed_files: List[str] = field(default_factory=list)
    total_time: float = 0.0
    
    def add_success(self) -> None:
        """添加成功计数"""
        self.successful += 1
        self.total_files += 1
    
    def add_failure(self, file_path: str) -> None:
        """添加失败计数和文件"""
        self.failed += 1
        self.total_files += 1
        self.failed_files.append(file_path)
