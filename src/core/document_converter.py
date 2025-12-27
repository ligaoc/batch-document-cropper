"""文档格式转换器 - 使用 LibreOffice 进行转换"""
import os
import subprocess
import shutil
from typing import Optional

from .file_validator import needs_conversion, get_file_extension, is_doc


class ConversionError(Exception):
    """文档转换错误"""
    def __init__(self, file_path: str, reason: str):
        self.file_path = file_path
        self.reason = reason
        super().__init__(f"转换失败 {file_path}: {reason}")


class DocumentConverter:
    """文档格式转换器，使用 LibreOffice 进行转换"""
    
    # LibreOffice 可能的安装路径
    LIBREOFFICE_PATHS = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    
    def __init__(self, libreoffice_path: Optional[str] = None):
        """初始化转换器
        
        Args:
            libreoffice_path: LibreOffice 可执行文件路径，None 则自动检测
        """
        self._libreoffice_path = libreoffice_path or self._detect_libreoffice()
    
    def _detect_libreoffice(self) -> Optional[str]:
        """自动检测 LibreOffice 安装路径"""
        # 先尝试 PATH 中的 soffice
        soffice = shutil.which("soffice")
        if soffice:
            return soffice
        
        # 检查常见安装路径
        for path in self.LIBREOFFICE_PATHS:
            if os.path.isfile(path):
                return path
        
        return None
    
    @property
    def is_available(self) -> bool:
        """检查 LibreOffice 是否可用"""
        return self._libreoffice_path is not None
    
    def is_conversion_needed(self, file_path: str) -> bool:
        """判断文件是否需要转换
        
        Args:
            file_path: 文件路径
            
        Returns:
            是否需要转换 (DOC/DOCX 返回 True)
        """
        return needs_conversion(file_path)
    
    def is_doc_file(self, file_path: str) -> bool:
        """判断文件是否为 DOC 格式
        
        Args:
            file_path: 文件路径
            
        Returns:
            是否为 DOC 文件
        """
        return is_doc(file_path)
    
    def convert_doc_to_docx(self, input_path: str, output_dir: str) -> str:
        """将 DOC 转换为 DOCX
        
        Args:
            input_path: 输入 DOC 文件路径
            output_dir: 输出目录
            
        Returns:
            转换后的 DOCX 文件路径
            
        Raises:
            ConversionError: 转换失败时抛出
        """
        if not self.is_available:
            raise ConversionError(input_path, "LibreOffice 未安装或未找到")
        
        if not os.path.isfile(input_path):
            raise ConversionError(input_path, "输入文件不存在")
        
        if not os.path.isdir(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        # 构建 LibreOffice 命令 - 转换为 DOCX
        cmd = [
            self._libreoffice_path,
            "--headless",
            "--convert-to", "docx",
            "--outdir", output_dir,
            input_path,
        ]
        
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120,  # 2分钟超时
            )
            
            if result.returncode != 0:
                raise ConversionError(input_path, result.stderr or "转换命令执行失败")
            
            # 构建输出文件路径
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_path = os.path.join(output_dir, f"{base_name}.docx")
            
            if not os.path.isfile(output_path):
                raise ConversionError(input_path, "转换后的 DOCX 文件未生成")
            
            return output_path
            
        except subprocess.TimeoutExpired:
            raise ConversionError(input_path, "转换超时")
        except FileNotFoundError:
            raise ConversionError(input_path, "LibreOffice 可执行文件未找到")
    
    def convert_to_pdf(self, input_path: str, output_dir: str) -> str:
        """将 DOC/DOCX 转换为 PDF (用于预览)
        
        Args:
            input_path: 输入文件路径
            output_dir: 输出目录
            
        Returns:
            转换后的 PDF 文件路径
            
        Raises:
            ConversionError: 转换失败时抛出
        """
        if not self.is_available:
            raise ConversionError(input_path, "LibreOffice 未安装或未找到")
        
        if not os.path.isfile(input_path):
            raise ConversionError(input_path, "输入文件不存在")
        
        if not os.path.isdir(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        # 构建 LibreOffice 命令
        cmd = [
            self._libreoffice_path,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", output_dir,
            input_path,
        ]
        
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120,  # 2分钟超时
            )
            
            if result.returncode != 0:
                raise ConversionError(input_path, result.stderr or "转换命令执行失败")
            
            # 构建输出文件路径
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_path = os.path.join(output_dir, f"{base_name}.pdf")
            
            if not os.path.isfile(output_path):
                raise ConversionError(input_path, "转换后的 PDF 文件未生成")
            
            return output_path
            
        except subprocess.TimeoutExpired:
            raise ConversionError(input_path, "转换超时")
        except FileNotFoundError:
            raise ConversionError(input_path, "LibreOffice 可执行文件未找到")
