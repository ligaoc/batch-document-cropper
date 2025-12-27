# Implementation Plan: Preview Button Loading

## Overview

本实现计划将预览功能从自动加载改为手动触发，添加加载动画，提升用户体验。主要修改 `preview_widget.py` 和 `main_window.py`。

## Tasks

- [x] 1. 创建加载指示器组件
  - [x] 1.1 实现 LoadingIndicator 类
    - 在 `src/gui/preview_widget.py` 中添加 LoadingIndicator 类
    - 使用 QLabel + QMovie 实现转圈动画
    - 实现 `start()` 和 `stop()` 方法
    - _Requirements: 2.1, 2.3, 2.4_

- [x] 2. 创建预览加载线程
  - [x] 2.1 实现 PreviewLoadingThread 类
    - 在 `src/gui/preview_widget.py` 中添加 PreviewLoadingThread 类
    - 继承 QThread
    - 实现 `loading_finished` 和 `loading_failed` 信号
    - 在 `run()` 方法中执行文档转换和预览生成
    - _Requirements: 2.1, 2.3, 2.4_

- [x] 3. 更新 PreviewWidget 组件
  - [x] 3.1 添加预览按钮和状态管理
    - 添加 PreviewState 枚举
    - 添加预览按钮 `_preview_btn`
    - 添加文件名标签 `_filename_label`
    - 添加错误信息标签 `_error_label`
    - 添加 LoadingIndicator 实例
    - _Requirements: 1.2, 1.4, 2.4_

  - [x] 3.2 修改 load_document 方法为 on_file_selected
    - 重命名方法为 `on_file_selected`
    - 只保存文件路径和边距设置
    - 显示文件名和预览按钮
    - 不自动加载预览
    - _Requirements: 1.1, 1.2_

  - [x] 3.3 实现 start_preview 方法
    - 显示加载动画
    - 禁用预览按钮
    - 启动 PreviewLoadingThread
    - _Requirements: 1.3, 2.1, 2.2_

  - [x] 3.4 实现加载完成和失败处理
    - 连接 PreviewLoadingThread 信号
    - 实现 `_on_loading_finished` 方法
    - 实现 `_on_loading_failed` 方法
    - _Requirements: 2.3, 2.4_

  - [x] 3.5 更新 update_margins 方法
    - 如果预览已加载，仅重绘裁剪线
    - 如果预览未加载，仅保存边距设置
    - 不触发预览加载
    - _Requirements: 3.1, 3.2, 3.3_

- [x] 4. 更新 MainWindow 集成
  - [x] 4.1 修改文件选择信号处理
    - 修改 `_on_file_selected` 方法
    - 调用 `_preview.on_file_selected()` 而不是 `load_document()`
    - _Requirements: 1.1_

  - [x] 4.2 修改边距变化信号处理
    - 确保 `_on_margins_changed` 调用 `_preview.update_margins()`
    - _Requirements: 3.1, 3.2_

- [x] 5. 检查点 - 确保功能正常
  - 手动测试：选择文件后显示预览按钮
  - 手动测试：点击预览按钮显示加载动画
  - 手动测试：加载完成后显示预览
  - 手动测试：修改边距只更新裁剪线
  - 如有问题，询问用户

- [ ]* 6. 编写属性测试
  - [ ]* 6.1 编写边距更新属性测试
    - **Property 1: Margin Update Without Document Reload**
    - **Validates: Requirements 3.1**

## Notes

- 标记 `*` 的任务为可选测试任务
- 主要修改文件：`src/gui/preview_widget.py` 和 `src/gui/main_window.py`
- 使用 PyQt5 的 QThread 实现后台加载
- 加载动画可使用内置的 busy indicator 或自定义 GIF

