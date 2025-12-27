"""核心处理模块"""
from .docx_cropper import DOCXCropper, DOCXCropError
from .pdf_cropper import PDFCropper
from .document_converter import DocumentConverter, ConversionError
from .batch_processor import BatchProcessor
from .resolution_keeper import ResolutionKeeper
from .output_manager import (
    OutputError,
    generate_output_filename,
    generate_output_path,
    validate_output_dir,
    ensure_output_dir
)
from .file_validator import is_supported_format, get_file_extension, needs_conversion
