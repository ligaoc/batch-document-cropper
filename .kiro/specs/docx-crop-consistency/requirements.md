# Requirements Document

## Introduction

修复 DOCX 文档裁剪行为与预览不一致的问题。当前预览显示的是从页面边缘裁剪掉指定毫米数的效果，但实际裁剪只是将页面边距设置为指定值，导致最终结果与预览不符。

## Glossary

- **Cropper**: 文档裁剪器，负责执行实际的裁剪操作
- **MarginSettings**: 边距设置数据类，包含上、下、左、右四个边距值（单位：毫米）
- **Preview_Widget**: 预览组件，显示裁剪效果预览
- **Original_Margin**: 文档原始边距，即裁剪前文档已有的页面边距
- **Crop_Amount**: 裁剪量，用户指定的从页面边缘裁剪掉的内容量

## Requirements

### Requirement 1: DOCX 裁剪量计算

**User Story:** As a user, I want the DOCX cropping to match the preview, so that the final document looks exactly as shown in the preview.

#### Acceptance Criteria

1. WHEN a user specifies a crop amount of X mm, THE Cropper SHALL calculate the new margin as: new_margin = original_margin + X mm
2. WHEN the original document has a top margin of 25mm and user specifies 28mm crop, THE Cropper SHALL set the final top margin to 53mm (25 + 28)
3. WHEN the crop amount is 0mm for any edge, THE Cropper SHALL preserve the original margin for that edge
4. THE Cropper SHALL apply the margin calculation consistently to all four edges (top, bottom, left, right)

### Requirement 2: 获取原始边距

**User Story:** As a developer, I want to retrieve the original margins from a DOCX document, so that I can calculate the correct new margins.

#### Acceptance Criteria

1. WHEN processing a DOCX file, THE Cropper SHALL read the original margins from the first section before applying changes
2. WHEN a section has no explicit margin defined, THE Cropper SHALL use the default Word margin value (typically 25.4mm or 1 inch)
3. WHEN a document has multiple sections with different margins, THE Cropper SHALL read and apply margins per-section

### Requirement 3: 预览与裁剪一致性

**User Story:** As a user, I want the preview to accurately represent the final cropped document, so that I can make informed decisions about crop settings.

#### Acceptance Criteria

1. WHEN the preview shows content being cropped at position X mm from the edge, THE Cropper SHALL produce output where that same content is cropped
2. IF the calculated new margin would exceed the page dimensions, THEN THE Cropper SHALL return an error indicating invalid crop settings
3. WHEN comparing preview and output, THE visual crop position SHALL match within 1mm tolerance

### Requirement 4: 边距验证

**User Story:** As a user, I want to be warned if my crop settings would result in invalid margins, so that I can adjust them before processing.

#### Acceptance Criteria

1. IF the new calculated margin (original + crop) exceeds half the page dimension, THEN THE Cropper SHALL return an error
2. WHEN validation fails, THE Cropper SHALL provide a clear error message indicating which edge has invalid settings
3. THE Cropper SHALL validate margins before making any changes to the document
