"""自定义异常类"""


class CropperError(Exception):
    """基础异常类"""
    pass


class FileFormatError(CropperError):
    """文件格式错误"""
    def __init__(self, file_path: str, extension: str):
        self.file_path = file_path
        self.extension = extension
        super().__init__(
            f"不支持的文件格式: {extension}。支持的格式: PDF, DOCX, DOC"
        )


class ConversionError(CropperError):
    """文档转换错误"""
    def __init__(self, file_path: str, reason: str):
        self.file_path = file_path
        self.reason = reason
        super().__init__(f"转换失败 {file_path}: {reason}")


class CropError(CropperError):
    """裁剪错误"""
    def __init__(self, file_path: str, reason: str):
        self.file_path = file_path
        self.reason = reason
        super().__init__(f"裁剪失败 {file_path}: {reason}")


class MarginValidationError(CropperError):
    """边距验证错误"""
    def __init__(self, field: str, value: float, reason: str):
        self.field = field
        self.value = value
        self.reason = reason
        super().__init__(f"无效的边距值 {field}: {value}。{reason}")


class OutputError(CropperError):
    """输出错误"""
    def __init__(self, output_path: str, reason: str):
        self.output_path = output_path
        self.reason = reason
        super().__init__(f"输出失败 {output_path}: {reason}")
