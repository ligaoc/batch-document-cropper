# Requirements Document

## Introduction

本文档定义了文档预览功能的用户体验改进需求。当前实现中，选中文件后会自动加载预览，但由于第三方工具（如 LibreOffice）转换速度较慢，导致用户体验不佳。本需求旨在通过添加手动预览按钮和加载动画来改善用户体验。

## Glossary

- **Preview_Widget**: 预览组件，负责显示文档预览和裁剪边界线
- **Loading_Indicator**: 加载指示器，显示预览加载中的动画效果
- **Preview_Button**: 预览按钮，用户点击后触发预览加载

## Requirements

### Requirement 1: 手动预览触发

**User Story:** As a user, I want to manually trigger document preview by clicking a button, so that I can control when the slow preview loading happens.

#### Acceptance Criteria

1. WHEN a user selects a file from the file list, THE Preview_Widget SHALL NOT automatically load the document preview
2. WHEN a user selects a file, THE Preview_Widget SHALL display the selected filename and a preview button
3. WHEN a user clicks the Preview_Button, THE Preview_Widget SHALL start loading the document preview
4. THE Preview_Button SHALL be clearly visible in the preview area

### Requirement 2: 加载状态指示

**User Story:** As a user, I want to see a loading indicator while the preview is being generated, so that I know the system is working.

#### Acceptance Criteria

1. WHEN the preview is loading, THE Loading_Indicator SHALL display a spinning animation
2. WHEN the preview is loading, THE Preview_Button SHALL be disabled to prevent multiple clicks
3. WHEN the preview loading completes successfully, THE Loading_Indicator SHALL stop and display the preview image
4. WHEN the preview loading fails, THE Loading_Indicator SHALL stop and display an error message

### Requirement 3: 边距变化时的预览更新

**User Story:** As a user, I want the preview to update when I change margin settings, so that I can see the effect of my changes.

#### Acceptance Criteria

1. WHEN margin settings are changed AND a preview is already loaded, THE Preview_Widget SHALL update the crop boundary lines without reloading the document
2. WHEN margin settings are changed AND no preview is loaded, THE Preview_Widget SHALL NOT trigger automatic preview loading
3. WHEN updating crop boundary lines, THE Preview_Widget SHALL NOT show the loading indicator

