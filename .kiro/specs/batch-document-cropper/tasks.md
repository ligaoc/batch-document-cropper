# Implementation Plan: Batch Document Cropper

## Overview

本实现计划将设计文档转化为可执行的编码任务。使用 Python 开发，PyQt5 构建 GUI，PyMuPDF 处理 PDF，python-docx 处理 DOCX，LibreOffice 进行 DOC 到 DOCX 转换，PyInstaller 打包为 EXE。

## Tasks

- [x] 1. 项目初始化和核心数据模型
  - [x] 1.1 创建项目结构和依赖配置
    - 创建 `src/` 目录结构
    - 创建 `requirements.txt` 包含 PyQt5, PyMuPDF, python-docx, hypothesis
    - 创建 `setup.py` 或 `pyproject.toml`
    - _Requirements: 8.1, 8.3_

  - [x] 1.2 实现 MarginSettings 数据类
    - 创建 `src/models/margin_settings.py`
    - 实现 top, bottom, left, right 属性
    - 实现 `validate()` 方法验证非负数值
    - 实现 `to_points()` 方法进行毫米到点数转换
    - _Requirements: 2.1, 2.2, 2.3, 2.4_

  - [ ]* 1.3 编写 MarginSettings 属性测试
    - **Property 2: Margin Settings Validation and Independence**
    - **Property 3: Margin Unit Conversion Round-Trip**
    - **Validates: Requirements 2.2, 2.3, 2.4**

  - [x] 1.4 实现任务和结果数据模型
    - 创建 `src/models/task.py`
    - 实现 ProcessingTask, TaskStatus, CropResult, ProcessingSummary
    - _Requirements: 4.2, 4.4_

- [x] 2. 检查点 - 确保数据模型测试通过
  - 运行所有测试，确保通过
  - 如有问题，询问用户

- [x] 3. 文件验证和文档转换模块
  - [x] 3.1 实现文件格式验证器
    - 创建 `src/core/file_validator.py`
    - 实现 `is_supported_format()` 函数
    - 实现 `get_file_extension()` 函数
    - 支持 PDF, DOCX, DOC 格式 (大小写不敏感)
    - _Requirements: 1.1, 1.3_

  - [ ]* 3.2 编写文件格式验证属性测试
    - **Property 1: File Format Validation**
    - **Validates: Requirements 1.1, 1.3**

  - [x] 3.3 实现 DocumentConverter 类 (DOC 到 DOCX 转换)
    - 创建 `src/core/document_converter.py`
    - 实现 LibreOffice 路径检测
    - 实现 `convert_doc_to_docx()` 方法
    - 实现 `is_doc_file()` 方法
    - _Requirements: 5.1, 5.2, 5.3, 5.4_

  - [ ]* 3.4 编写 DocumentConverter 单元测试
    - 测试 DOC 到 DOCX 转换
    - 测试 LibreOffice 不可用时的错误处理
    - **Property 14: DOC to DOCX Conversion Fidelity**
    - _Requirements: 5.1, 5.2, 5.3, 5.4_

- [x] 4. PDF 裁剪核心模块
  - [x] 4.1 实现 PDFCropper 类
    - 创建 `src/core/pdf_cropper.py`
    - 实现 `crop()` 方法使用 PyMuPDF 裁剪 PDF
    - 实现 `get_page_count()` 方法
    - 实现 `generate_preview()` 方法生成预览图像
    - _Requirements: 11.1, 11.3_

  - [ ]* 4.2 编写 PDFCropper 属性测试
    - **Property 9: Cropping Accuracy Across Page Sizes**
    - **Property 10: Processing Idempotence**
    - **Validates: Requirements 11.1, 11.2, 11.3**

  - [x] 4.3 实现 ResolutionKeeper 类
    - 创建 `src/core/resolution_keeper.py`
    - 实现 `get_pdf_resolution()` 方法
    - 实现 `verify_resolution()` 方法
    - _Requirements: 6.1, 6.2, 6.4_

  - [ ]* 4.4 编写分辨率保持属性测试
    - **Property 6: Resolution Preservation Invariant**
    - **Validates: Requirements 6.1, 6.2, 6.4**

- [x] 5. DOCX 裁剪核心模块 (新增)
  - [x] 5.1 实现 DOCXCropper 类
    - 创建 `src/core/docx_cropper.py`
    - 使用 python-docx 库
    - 实现 `crop()` 方法通过调整页面边距裁剪 DOCX
    - 实现 `get_section_count()` 方法获取节数
    - 确保所有节的边距都被调整
    - _Requirements: 10.1, 10.2, 10.3, 10.4_

  - [ ]* 5.2 编写 DOCXCropper 属性测试
    - **Property 12: DOCX Margin Adjustment Accuracy**
    - **Property 13: DOCX Content Preservation**
    - **Validates: Requirements 10.1, 10.2, 10.3, 10.4**

- [x] 6. 检查点 - 确保核心模块测试通过
  - 运行所有测试，确保通过
  - 如有问题，询问用户

- [x] 7. 批量处理模块 (更新)
  - [x] 7.1 更新 BatchProcessor 类以支持格式保持
    - 修改 `src/core/batch_processor.py`
    - PDF 文件: 使用 PDFCropper，输出 PDF
    - DOCX 文件: 使用 DOCXCropper，输出 DOCX
    - DOC 文件: 先转换为 DOCX，再使用 DOCXCropper，输出 DOCX
    - 实现任务队列管理 (最多 5 个并发)
    - _Requirements: 1.2, 4.1, 4.2, 4.3, 4.4, 7.4, 7.5, 7.6, 7.7_

  - [ ]* 7.2 编写 BatchProcessor 属性测试
    - **Property 5: Batch Processing with Failure Isolation**
    - **Property 7: Output Format Preservation**
    - **Property 11: Queue Size Limit Enforcement**
    - **Validates: Requirements 1.2, 4.1, 4.3, 4.4, 7.4, 7.5, 7.6, 7.7**

  - [x] 7.3 更新输出文件命名逻辑
    - 修改 `src/core/output_manager.py`
    - PDF 输入: 原名 + 后缀 + .pdf
    - DOCX 输入: 原名 + 后缀 + .docx
    - DOC 输入: 原名 + 后缀 + .docx
    - 实现输出目录验证
    - _Requirements: 7.2, 7.3, 7.4, 7.5, 7.6_

  - [ ]* 7.4 编写输出文件名属性测试
    - **Property 8: Output Filename Pattern Preservation**
    - **Validates: Requirements 7.2**

- [ ] 8. 检查点 - 确保批处理模块测试通过
  - 运行所有测试，确保通过
  - 如有问题，询问用户

- [x] 9. GUI 界面实现
  - [x] 9.1 实现主窗口框架
    - 创建 `src/gui/main_window.py`
    - 实现 MainWindow 类继承 QMainWindow
    - 设置窗口标题、大小、布局
    - _Requirements: 8.4_

  - [x] 9.2 实现文件列表组件
    - 创建 `src/gui/file_list_widget.py`
    - 实现文件拖放添加
    - 实现文件选择对话框
    - 实现文件列表显示和删除
    - _Requirements: 1.1, 1.2_

  - [x] 9.3 实现边距设置面板
    - 创建 `src/gui/margin_panel.py`
    - 实现四个边距输入框 (QSpinBox)
    - 实现实时验证和错误提示
    - _Requirements: 2.1, 2.2, 2.3, 2.4_

  - [x] 9.4 实现预览组件
    - 创建 `src/gui/preview_widget.py`
    - 实现文档预览显示
    - 实现裁剪边界线绘制
    - 实现多页导航
    - _Requirements: 3.1, 3.2, 3.3, 3.4_

  - [ ]* 9.5 编写预览组件属性测试
    - **Property 4: Preview Generation Consistency**
    - **Validates: Requirements 3.1, 3.2, 3.4**

  - [x] 9.6 实现进度显示组件
    - 创建 `src/gui/progress_widget.py`
    - 实现每个文件的进度条
    - 实现处理状态显示
    - _Requirements: 4.2_

  - [x] 9.7 集成所有 GUI 组件
    - 在 MainWindow 中组装所有组件
    - 连接信号和槽
    - 实现开始处理和取消按钮
    - _Requirements: 8.4_

- [x] 10. 检查点 - 确保 GUI 功能正常
  - 手动测试 GUI 交互
  - 如有问题，询问用户

- [x] 11. 异常处理和错误提示
  - [x] 11.1 实现自定义异常类
    - 创建 `src/exceptions.py`
    - 实现 FileFormatError, ConversionError, CropError, MarginValidationError, OutputError
    - _Requirements: 1.3, 5.5, 7.3_

  - [x] 11.2 集成错误处理到各模块
    - 在 DocumentConverter 中添加转换错误处理
    - 在 PDFCropper 中添加裁剪错误处理
    - 在 DOCXCropper 中添加裁剪错误处理
    - 在 BatchProcessor 中添加任务失败处理
    - 在 GUI 中添加错误对话框显示
    - _Requirements: 4.3, 5.5, 7.3_

- [x] 12. 应用入口和打包配置
  - [x] 12.1 创建应用入口
    - 创建 `src/main.py`
    - 实现 QApplication 初始化
    - 实现主窗口启动
    - _Requirements: 8.4_

  - [x] 12.2 创建 PyInstaller 打包配置
    - 创建 `build.spec` 文件
    - 配置单文件 EXE 打包
    - 配置图标和元数据
    - 配置 LibreOffice 依赖处理
    - _Requirements: 8.1, 8.2, 8.3_

- [ ] 13. 最终检查点 - 确保所有测试通过
  - 运行完整测试套件
  - 验证 EXE 打包成功
  - 如有问题，询问用户

## Notes

- 标记 `*` 的任务为可选测试任务，可跳过以加快 MVP 开发
- 每个任务引用具体需求以确保可追溯性
- 检查点确保增量验证
- 属性测试验证通用正确性属性
- 单元测试验证特定示例和边界情况
- **重要变更**: 输出格式保持与输入格式一致 (PDF→PDF, DOCX→DOCX, DOC→DOCX)
