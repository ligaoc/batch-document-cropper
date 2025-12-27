# 批量文档裁剪工具

一款基于 Python 的批量文档裁剪桌面工具，支持 PDF、DOCX、DOC 格式文件的批量裁剪处理。

## 功能特性

- 支持 PDF、DOCX、DOC 文件格式
- 自定义上下左右边距裁剪
- 实时预览裁剪效果
- 批量处理最多 5 个文档
- 保持原始分辨率和格式
- 可打包为独立 EXE 可执行文件

## 系统要求

- Python 3.9+
- LibreOffice (用于 DOC/DOCX 转换)
- Windows 操作系统

## 安装

1. 安装依赖:

```bash
pip install -r requirements.txt
```

2. 安装 LibreOffice (用于 Word 文档转换):
   - 下载地址: https://www.libreoffice.org/download/download/

## 运行

```bash
python run.py
```

## 打包为 EXE

```bash
pyinstaller build.spec
```

打包后的 EXE 文件位于 `dist/BatchDocumentCropper.exe`

## 使用说明

1. 点击"添加文件"选择要处理的文档 (最多 5 个)
2. 设置上下左右边距值 (单位: 毫米)
3. 选择文件可预览裁剪效果
4. 点击"选择目录"设置输出位置
5. 点击"开始处理"执行批量裁剪

## 项目结构

```
├── src/
│   ├── core/           # 核心处理模块
│   │   ├── batch_processor.py
│   │   ├── document_converter.py
│   │   ├── file_validator.py
│   │   ├── output_manager.py
│   │   ├── pdf_cropper.py
│   │   └── resolution_keeper.py
│   ├── gui/            # GUI 界面模块
│   │   ├── file_list_widget.py
│   │   ├── main_window.py
│   │   ├── margin_panel.py
│   │   ├── preview_widget.py
│   │   └── progress_widget.py
│   ├── models/         # 数据模型
│   │   ├── margin_settings.py
│   │   └── task.py
│   ├── exceptions.py   # 自定义异常
│   └── main.py         # 应用入口
├── tests/              # 测试
├── build.spec          # PyInstaller 配置
├── requirements.txt    # 依赖
└── run.py              # 启动脚本
```

## 许可证

MIT License
