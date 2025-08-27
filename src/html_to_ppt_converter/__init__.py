"""
HTML到PPT转换器
独立的HTML幻灯片到PowerPoint转换工具
"""

__version__ = "1.0.0"
__author__ = "LandPPT Team"
__description__ = "HTML到PPT转换器 - 将HTML幻灯片转换为PowerPoint文件"

from .config import get_config, load_config
from .pdf_converter import get_pdf_converter, convert_html_to_pdf
from .pptx_converter import get_pptx_converter, convert_pdf_to_pptx

__all__ = [
    'get_config',
    'load_config', 
    'get_pdf_converter',
    'convert_html_to_pdf',
    'get_pptx_converter',
    'convert_pdf_to_pptx'
]
