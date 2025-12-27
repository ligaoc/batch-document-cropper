"""文档格式转换器 - 优先使用 Microsoft Word，备选 LibreOffice"""
import os
import subprocess
import shutil
import sys
import logging
import time
from typing import Optional

from .file_validator import needs_conversion, get_file_extension, is_doc

# 配置日志
logger = logging.getLogger(__name__)

# 检查 docx2pdf 是否可用（需要 Microsoft Word）
_DOCX2PDF_AVAILABLE = False
try:
    if sys.platform == 'win32':
        import docx2pdf
        _DOCX2PDF_AVAILABLE = True
        logger.info("docx2pdf 模块已加载")
except ImportError:
    logger.warning("docx2pdf 模块未安装")

# 全局缓存 Word 可用性检测结果，避免重复检测
_WORD_AVAILABLE_CACHE = None

# 全局单例实例
_CONVERTER_INSTANCE = None


def get_document_converter() -> 'DocumentConverter':
    """获取全局唯一的 DocumentConverter 实例（单例模式）"""
    global _CONVERTER_INSTANCE
    if _CONVERTER_INSTANCE is None:
        _CONVERTER_INSTANCE = DocumentConverter()
    return _CONVERTER_INSTANCE


class ConversionError(Exception):
    """文档转换错误"""
    def __init__(self, file_path: str, reason: str):
        self.file_path = file_path
        self.reason = reason
        super().__init__(f"转换失败 {file_path}: {reason}")


class DocumentConverter:
    """文档格式转换器
    
    优先使用 Microsoft Word（通过 docx2pdf），因为它能保证 100% 准确的分页和格式。
    如果 Word 不可用，则回退到 LibreOffice。
    """
    
    # LibreOffice 可能的安装路径
    LIBREOFFICE_PATHS = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    
    def __init__(self, libreoffice_path: Optional[str] = None, prefer_word: bool = True):
        """初始化转换器
        
        Args:
            libreoffice_path: LibreOffice 可执行文件路径，None 则自动检测
            prefer_word: 是否优先使用 Microsoft Word（默认 True）
        """
        self._libreoffice_path = libreoffice_path or self._detect_libreoffice()
        self._prefer_word = prefer_word and _DOCX2PDF_AVAILABLE
        self._word_available = self._check_word_available()
        
        # 打印初始化信息
        logger.info(f"=== 文档转换器初始化 ===")
        logger.info(f"docx2pdf 可用: {_DOCX2PDF_AVAILABLE}")
        logger.info(f"Word 可用: {self._word_available}")
        logger.info(f"LibreOffice 路径: {self._libreoffice_path}")
        logger.info(f"优先使用 Word: {self._prefer_word}")
        logger.info(f"当前转换器: {self.converter_name}")
        print(f"[转换器] 初始化完成，使用: {self.converter_name}")
        
        # 如果配置为使用 Word 但 Word 不可用，直接报错
        if self._prefer_word and not self._word_available:
            error_msg = "Microsoft Word 未安装或不可用，请安装 Microsoft Office 后重试"
            logger.error(error_msg)
            raise RuntimeError(error_msg)
    
    def _check_word_available(self) -> bool:
        """检查 Microsoft Word 是否可用（使用缓存避免重复检测）"""
        global _WORD_AVAILABLE_CACHE
        
        # 如果已经检测过，直接返回缓存结果
        if _WORD_AVAILABLE_CACHE is not None:
            logger.info(f"使用缓存的 Word 检测结果: {_WORD_AVAILABLE_CACHE}")
            return _WORD_AVAILABLE_CACHE
        
        if not _DOCX2PDF_AVAILABLE:
            logger.info("docx2pdf 不可用，跳过 Word 检测")
            _WORD_AVAILABLE_CACHE = False
            return False
        
        try:
            # 尝试创建 Word COM 对象来检测 Word 是否安装
            import win32com.client
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.Quit()
            # 释放 COM 对象并等待一下
            del word
            time.sleep(0.5)  # 等待 Word 进程完全退出
            logger.info("Microsoft Word COM 对象创建成功")
            _WORD_AVAILABLE_CACHE = True
            return True
        except Exception as e:
            logger.warning(f"Microsoft Word 不可用: {e}")
            _WORD_AVAILABLE_CACHE = False
            return False
    
    @property
    def word_available(self) -> bool:
        """检查 Microsoft Word 是否可用"""
        return self._word_available
    
    @property
    def converter_name(self) -> str:
        """获取当前使用的转换器名称"""
        if self._prefer_word and self._word_available:
            return "Microsoft Word"
        elif self._libreoffice_path:
            return "LibreOffice"
        return "无可用转换器"
    
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
        """检查是否有可用的转换器（Word 或 LibreOffice）"""
        return self._word_available or self._libreoffice_path is not None
    
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
        
        优先使用 Microsoft Word 转换（100% 准确），如果不可用则使用 LibreOffice。
        
        Args:
            input_path: 输入文件路径
            output_dir: 输出目录
            
        Returns:
            转换后的 PDF 文件路径
            
        Raises:
            ConversionError: 转换失败时抛出
        """
        logger.info(f"=== 开始转换 ===")
        logger.info(f"输入文件: {input_path}")
        logger.info(f"输出目录: {output_dir}")
        print(f"[转换器] 开始转换: {os.path.basename(input_path)}")
        
        if not self.is_available:
            raise ConversionError(input_path, "没有可用的转换器（需要 Microsoft Word 或 LibreOffice）")
        
        if not os.path.isfile(input_path):
            raise ConversionError(input_path, "输入文件不存在")
        
        if not os.path.isdir(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        # 优先使用 Microsoft Word（通过 docx2pdf）
        if self._prefer_word and self._word_available:
            logger.info(">>> 使用 Microsoft Word 转换")
            print(f"[转换器] 使用 Microsoft Word 进行转换")
            return self._convert_to_pdf_with_word(input_path, output_dir)
        
        # 回退到 LibreOffice
        logger.info(">>> 使用 LibreOffice 转换")
        print(f"[转换器] 使用 LibreOffice 进行转换")
        return self._convert_to_pdf_with_libreoffice(input_path, output_dir)
    
    def _convert_to_pdf_with_word(self, input_path: str, output_dir: str) -> str:
        """使用 Microsoft Word 转换为 PDF（最准确）"""
        import pythoncom
        
        # 构建输出文件路径
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        output_path = os.path.join(output_dir, f"{base_name}.pdf")
        
        # 转换为绝对路径
        input_path = os.path.abspath(input_path)
        output_path = os.path.abspath(output_path)
        
        logger.info(f"使用 win32com 直接调用 Word")
        logger.info(f"输入: {input_path}")
        logger.info(f"输出: {output_path}")
        print(f"[Word] 正在转换...")
        
        word = None
        doc = None
        try:
            # 在当前线程初始化 COM
            pythoncom.CoInitialize()
            
            import win32com.client
            
            # 创建 Word 应用
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            
            # 打开文档
            doc = word.Documents.Open(input_path, ReadOnly=True)
            
            # 保存为 PDF (FileFormat=17 是 PDF 格式)
            doc.SaveAs2(output_path, FileFormat=17)
            
            print(f"[Word] 转换完成: {output_path}")
            
        except Exception as e:
            logger.error(f"Word 转换过程出错: {e}")
            # 检查文件是否已经生成
            if os.path.isfile(output_path) and os.path.getsize(output_path) > 0:
                print(f"[Word] 虽然出错但 PDF 已生成，继续使用")
                logger.info(f"PDF 文件已生成，忽略关闭时的错误")
            else:
                print(f"[Word] 转换失败: {e}")
                raise ConversionError(input_path, f"Word 转换失败: {e}")
        finally:
            # 确保关闭文档和 Word（忽略关闭时的错误）
            try:
                if doc:
                    doc.Close(False)
            except Exception:
                pass
            try:
                if word:
                    word.Quit()
            except Exception:
                pass
            try:
                # 释放 COM
                pythoncom.CoUninitialize()
            except Exception:
                pass
            # 等待一下让 Word 进程完全退出
            time.sleep(0.3)
        
        if not os.path.isfile(output_path):
            raise ConversionError(input_path, "转换后的 PDF 文件未生成")
        
        logger.info(f"Word 转换成功: {output_path}")
        return output_path
    
    def _convert_to_pdf_with_libreoffice(self, input_path: str, output_dir: str) -> str:
        """使用 LibreOffice 转换为 PDF"""
        if not self._libreoffice_path:
            raise ConversionError(input_path, "LibreOffice 未安装或未找到")
        
        logger.info(f"使用 LibreOffice: {self._libreoffice_path}")
        print(f"[LibreOffice] 正在转换...")
        
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
                logger.error(f"LibreOffice 转换失败: {result.stderr}")
                raise ConversionError(input_path, result.stderr or "转换命令执行失败")
            
            # 构建输出文件路径
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_path = os.path.join(output_dir, f"{base_name}.pdf")
            
            if not os.path.isfile(output_path):
                raise ConversionError(input_path, "转换后的 PDF 文件未生成")
            
            logger.info(f"LibreOffice 转换成功: {output_path}")
            print(f"[LibreOffice] 转换完成: {output_path}")
            return output_path
            
        except subprocess.TimeoutExpired:
            raise ConversionError(input_path, "转换超时")
        except FileNotFoundError:
            raise ConversionError(input_path, "LibreOffice 可执行文件未找到")
