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
- WPS Office (用于 DOC/DOCX 转 PDF)
- Windows 操作系统

## 安装

### 1. 安装 Python 依赖

```bash
pip install -r requirements.txt
```

### 2. 安装 WPS Office

WPS Office 用于 Word 文档 (DOC/DOCX) 转换为 PDF。

- 下载地址: https://www.wps.cn/
- 安装后无需额外配置，程序会自动调用 WPS 进行转换

## 运行

### 开发环境运行

```bash
python run.py
```

## 打包为 EXE

### 打包步骤

1. 确保已安装 PyInstaller:

```bash
pip install pyinstaller
```

2. 执行打包命令:

```bash
python -m PyInstaller build.spec --clean
```

3. 打包完成后，EXE 文件位于 `dist/BatchDocumentCropper.exe`

### 打包说明

- `--clean` 参数会清除之前的缓存，确保使用最新配置
- 打包后的 EXE 是单文件模式，可独立运行
- 运行 EXE 的电脑也需要安装 WPS Office 才能处理 Word 文档

## 部署

### 直接运行 EXE

1. 将 `dist/BatchDocumentCropper.exe` 复制到目标电脑
2. 确保目标电脑已安装 WPS Office
3. 双击运行即可

### 注意事项

- 首次运行可能需要几秒钟加载时间
- 如果杀毒软件误报，请添加信任
- 处理大文件时请耐心等待

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
│   │   ├── docx_cropper.py
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
├── build.spec          # PyInstaller 打包配置
├── requirements.txt    # Python 依赖
└── run.py              # 启动脚本
```

## 常见问题

### Q: 打包后运行报错 "pyinstaller 不是内部或外部命令"
A: 使用 `python -m PyInstaller build.spec` 代替直接运行 `pyinstaller`

### Q: Word 文档无法转换
A: 请确保已安装 WPS Office，并且 WPS 可以正常打开该文档

### Q: EXE 运行时闪退
A: 可以在命令行中运行 EXE 查看错误信息，或将 `build.spec` 中的 `console=False` 改为 `console=True` 重新打包

## 许可证

MIT License
