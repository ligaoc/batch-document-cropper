# Design Document: DOCX Crop Consistency

## Overview

修复 DOCX 文档裁剪行为与预览不一致的问题。核心改动是将裁剪逻辑从"设置边距"改为"增加边距"，即：新边距 = 原始边距 + 裁剪量。

## Architecture

当前架构保持不变，只需修改 `DOCXCropper.crop()` 方法的边距计算逻辑：

```
用户输入裁剪量 → 读取原始边距 → 计算新边距 → 验证有效性 → 应用到文档
```

## Components and Interfaces

### DOCXCropper 类修改

```python
class DOCXCropper:
    """DOCX 裁剪器，通过调整页面边距实现裁剪"""
    
    # Word 默认边距 (1 inch = 25.4mm)
    DEFAULT_MARGIN_MM = 25.4
    
    def crop(self, input_path: str, output_path: str,
             margins: MarginSettings) -> CropResult:
        """裁剪 DOCX 文件
        
        Args:
            input_path: 输入 DOCX 路径
            output_path: 输出 DOCX 路径
            margins: 裁剪量设置（从页面边缘裁剪掉的毫米数）
            
        Returns:
            裁剪结果对象
        """
        # 1. 验证输入
        # 2. 打开文档
        # 3. 对每个节：
        #    a. 读取原始边距
        #    b. 计算新边距 = 原始边距 + 裁剪量
        #    c. 验证新边距不超过页面尺寸的一半
        #    d. 应用新边距
        # 4. 保存文档
    
    def _get_section_margin(self, section, edge: str) -> float:
        """获取节的指定边距（毫米）
        
        Args:
            section: docx Section 对象
            edge: 边缘名称 ('top', 'bottom', 'left', 'right')
            
        Returns:
            边距值（毫米），如果未定义则返回默认值
        """
        margin_attr = f"{edge}_margin"
        margin = getattr(section, margin_attr, None)
        if margin is None:
            return self.DEFAULT_MARGIN_MM
        return margin.mm
    
    def _validate_new_margins(self, section, new_margins: MarginSettings) -> tuple[bool, str]:
        """验证新边距是否有效
        
        Args:
            section: docx Section 对象
            new_margins: 计算后的新边距
            
        Returns:
            (是否有效, 错误信息)
        """
        page_width = section.page_width.mm if section.page_width else 210  # A4 默认
        page_height = section.page_height.mm if section.page_height else 297
        
        if new_margins.left + new_margins.right >= page_width:
            return False, f"左右边距之和 ({new_margins.left + new_margins.right}mm) 超过页面宽度 ({page_width}mm)"
        if new_margins.top + new_margins.bottom >= page_height:
            return False, f"上下边距之和 ({new_margins.top + new_margins.bottom}mm) 超过页面高度 ({page_height}mm)"
        
        return True, ""
```

## Data Models

无需修改现有数据模型。`MarginSettings` 的语义从"目标边距"变为"裁剪量"，但数据结构保持不变。

## Correctness Properties

*A property is a characteristic or behavior that should hold true across all valid executions of a system-essentially, a formal statement about what the system should do.*

### Property 1: 边距累加正确性

*For any* DOCX 文档的任意节，和任意非负裁剪量设置（包括 0），裁剪后该节的边距应等于原始边距加上裁剪量。当裁剪量为 0 时，边距保持不变。

**Validates: Requirements 1.1, 1.2, 1.3, 1.4**

### Property 2: 边距验证与文档保护

*For any* DOCX 文档和裁剪量设置，如果计算后的左右边距之和超过页面宽度，或上下边距之和超过页面高度，裁剪操作应返回错误且不修改原文档。

**Validates: Requirements 3.2, 4.1, 4.3**

### Property 3: 原始边距读取正确性

*For any* DOCX 文档，裁剪器读取的原始边距应与文档实际边距一致。对于未定义边距的节，应使用默认值 25.4mm。

**Validates: Requirements 2.1, 2.2**

## Error Handling

| 错误场景 | 处理方式 |
|---------|---------|
| 输入文件不存在 | 返回 `CropResult.failure_result` |
| 裁剪量为负数 | 返回验证错误 |
| 新边距超过页面尺寸 | 返回错误，说明哪个边距无效 |
| 文档无法打开 | 返回异常信息 |

## Testing Strategy

### 单元测试

1. 测试边距累加计算
2. 测试零裁剪保持原样
3. 测试边距验证逻辑
4. 测试多节文档处理

### 属性测试

使用 `hypothesis` 库进行属性测试：

1. **Property 1**: 生成随机原始边距和裁剪量，验证输出边距 = 原始 + 裁剪量
2. **Property 2**: 生成随机文档，裁剪量全为 0，验证边距不变
3. **Property 3**: 生成会导致边距超限的裁剪量，验证返回错误
