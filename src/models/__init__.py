"""数据模型模块"""
from .margin_settings import MarginSettings
from .task import ProcessingTask, TaskStatus, CropResult, ProcessingSummary

__all__ = [
    "MarginSettings",
    "ProcessingTask",
    "TaskStatus", 
    "CropResult",
    "ProcessingSummary",
]
