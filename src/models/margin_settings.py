"""边距设置数据类"""
from dataclasses import dataclass
from typing import Tuple

# 1 inch = 25.4 mm, 1 inch = 72 points
MM_TO_POINTS = 72 / 25.4


@dataclass
class MarginSettings:
    """边距设置数据类
    
    Attributes:
        top: 上边距 (mm)
        bottom: 下边距 (mm)
        left: 左边距 (mm)
        right: 右边距 (mm)
    """
    top: float = 0.0
    bottom: float = 0.0
    left: float = 0.0
    right: float = 0.0
    
    def validate(self) -> Tuple[bool, str]:
        """验证边距设置是否有效
        
        Returns:
            (是否有效, 错误信息)
        """
        if self.top < 0:
            return False, "上边距不能为负数"
        if self.bottom < 0:
            return False, "下边距不能为负数"
        if self.left < 0:
            return False, "左边距不能为负数"
        if self.right < 0:
            return False, "右边距不能为负数"
        return True, ""
    
    def to_points(self) -> Tuple[float, float, float, float]:
        """将毫米转换为 PDF 点数 (1 inch = 72 points, 1 inch = 25.4 mm)
        
        Returns:
            (top_pt, bottom_pt, left_pt, right_pt)
        """
        return (
            self.top * MM_TO_POINTS,
            self.bottom * MM_TO_POINTS,
            self.left * MM_TO_POINTS,
            self.right * MM_TO_POINTS,
        )
    
    @staticmethod
    def from_points(top_pt: float, bottom_pt: float, 
                    left_pt: float, right_pt: float) -> "MarginSettings":
        """从点数创建 MarginSettings
        
        Args:
            top_pt: 上边距 (points)
            bottom_pt: 下边距 (points)
            left_pt: 左边距 (points)
            right_pt: 右边距 (points)
            
        Returns:
            MarginSettings 实例
        """
        return MarginSettings(
            top=top_pt / MM_TO_POINTS,
            bottom=bottom_pt / MM_TO_POINTS,
            left=left_pt / MM_TO_POINTS,
            right=right_pt / MM_TO_POINTS,
        )
