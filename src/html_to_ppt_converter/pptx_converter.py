#!/usr/bin/env python3
"""
简化的PDF到PPTX转换器
基于Apryse SDK，专门用于PPT转换
"""

import os
import tempfile
import logging
from pathlib import Path
from typing import Tuple, Optional
import platform

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class PDFToPPTXConverter:
    """PDF到PPTX转换器"""
    
    def __init__(self, license_key: Optional[str] = None):
        self.license_key = license_key
        self._sdk_initialized = False
        self._sdk_available = None
    
    def is_available(self) -> bool:
        """检查Apryse SDK是否可用"""
        if self._sdk_available is not None:
            return self._sdk_available
        
        try:
            import apryse_sdk
            self._sdk_available = True
            logger.info("Apryse SDK可用")
        except ImportError:
            self._sdk_available = False
            logger.warning("Apryse SDK不可用，请安装Apryse SDK")
        
        return self._sdk_available
    
    def _initialize_sdk(self) -> bool:
        """初始化Apryse SDK"""
        if self._sdk_initialized:
            return True

        if not self.is_available():
            return False

        try:
            from apryse_sdk.PDFNetPython import PDFNet

            # 首先尝试使用提供的许可证密钥
            if self.license_key:
                try:
                    PDFNet.Initialize(self.license_key)
                    logger.info("Apryse SDK初始化成功（使用许可证密钥）")
                except Exception as license_error:
                    logger.warning(f"许可证密钥无效: {license_error}")
                    logger.info("尝试使用试用模式...")
                    PDFNet.Initialize()
                    logger.info("Apryse SDK初始化成功（试用模式 - 输出将包含水印）")
            else:
                PDFNet.Initialize()
                logger.info("Apryse SDK初始化成功（试用模式 - 输出将包含水印）")

            # 添加资源搜索路径
            self._setup_resource_paths()

            self._sdk_initialized = True
            return True

        except Exception as e:
            logger.error(f"Apryse SDK初始化失败: {e}")
            return False

    def _setup_resource_paths(self):
        """设置资源搜索路径"""
        try:
            from apryse_sdk.PDFNetPython import PDFNet

            # 可能的资源路径
            resource_paths = [
                "./StructuredOutputModuleWindows",  # 解压后的模块目录
                "./StructuredOutputModuleWindows/Lib",  # 模块的Lib目录
                "./StructuredOutputModule",  # 备用名称
                "./Lib",  # Lib目录
                "./Resources",  # Resources目录
                os.path.join(os.path.dirname(__file__), "StructuredOutputModuleWindows"),
                os.path.join(os.path.dirname(__file__), "StructuredOutputModuleWindows", "Lib"),
                os.path.join(os.path.dirname(__file__), "StructuredOutputModule"),
                os.path.join(os.path.dirname(__file__), "Lib"),
                os.path.join(os.path.dirname(__file__), "Resources"),
            ]

            added_paths = []
            for path in resource_paths:
                if os.path.exists(path):
                    abs_path = os.path.abspath(path)
                    logger.info(f"添加资源搜索路径: {abs_path}")
                    PDFNet.AddResourceSearchPath(abs_path)
                    added_paths.append(abs_path)

            if not added_paths:
                logger.warning("未找到任何有效的资源路径")
            else:
                logger.info(f"总共添加了 {len(added_paths)} 个资源路径")

        except Exception as e:
            logger.warning(f"设置资源路径时出现警告: {e}")
    
    def convert_pdf_to_pptx(self, pdf_path: str, output_path: Optional[str] = None) -> Tuple[bool, str]:
        """
        将PDF文件转换为PPTX格式
        
        Args:
            pdf_path: PDF文件路径
            output_path: PPTX输出路径（可选，会自动生成）
        
        Returns:
            Tuple[bool, str]: (成功状态, 输出路径或错误信息)
        """
        if not self._initialize_sdk():
            error_msg = "Apryse SDK不可用，请检查安装和许可证"
            logger.error(error_msg)
            return False, error_msg
        
        # 验证输入文件
        if not os.path.exists(pdf_path):
            error_msg = f"PDF文件不存在: {pdf_path}"
            logger.error(error_msg)
            return False, error_msg
        
        # 生成输出路径
        if output_path is None:
            pdf_name = Path(pdf_path).stem
            output_path = str(Path(pdf_path).parent / f"{pdf_name}.pptx")
        
        try:
            from apryse_sdk.PDFNetPython import Convert
            
            logger.info(f"开始PDF到PPTX转换: {pdf_path} -> {output_path}")
            
            # 执行转换
            Convert.ToPowerPoint(pdf_path, output_path)
            
            # 验证输出文件
            if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                logger.info(f"PDF到PPTX转换成功: {output_path}")
                return True, output_path
            else:
                error_msg = "转换完成但输出文件为空或未创建"
                logger.error(error_msg)
                return False, error_msg
                
        except Exception as e:
            error_msg = f"PDF到PPTX转换失败: {str(e)}"
            logger.error(error_msg)
            return False, error_msg
    
    def convert_with_temp_pdf(self, pdf_content: bytes, output_filename: str = "converted.pptx") -> Tuple[bool, str, str]:
        """
        使用临时文件转换PDF内容为PPTX
        
        Args:
            pdf_content: PDF文件内容（字节）
            output_filename: 输出文件名
        
        Returns:
            Tuple[bool, str, str]: (成功状态, 输出路径, 错误信息)
        """
        temp_pdf_path = None
        temp_pptx_path = None
        
        try:
            # 创建临时PDF文件
            with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_pdf:
                temp_pdf.write(pdf_content)
                temp_pdf_path = temp_pdf.name
            
            # 创建临时PPTX文件路径
            temp_dir = tempfile.gettempdir()
            base_name = Path(output_filename).stem
            temp_pptx_path = os.path.join(temp_dir, f"{base_name}_{os.getpid()}.pptx")
            
            # 执行转换
            success, result = self.convert_pdf_to_pptx(temp_pdf_path, temp_pptx_path)
            
            if success:
                return True, temp_pptx_path, ""
            else:
                return False, "", result
                
        except Exception as e:
            error_msg = f"临时文件转换失败: {str(e)}"
            logger.error(error_msg)
            return False, "", error_msg
        finally:
            # 清理临时PDF文件
            if temp_pdf_path and os.path.exists(temp_pdf_path):
                try:
                    os.unlink(temp_pdf_path)
                except:
                    pass
    
    def get_sdk_info(self) -> dict:
        """获取SDK信息"""
        info = {
            "available": self.is_available(),
            "initialized": self._sdk_initialized,
            "platform": platform.system(),
            "license_provided": bool(self.license_key)
        }
        
        if self.is_available():
            try:
                from apryse_sdk.PDFNetPython import PDFNet
                info["version"] = PDFNet.GetVersion()
            except:
                info["version"] = "Unknown"
        
        return info


# 全局转换器实例
_pptx_converter = None

def get_pptx_converter(license_key: Optional[str] = None) -> PDFToPPTXConverter:
    """获取PPTX转换器实例"""
    global _pptx_converter
    if _pptx_converter is None:
        _pptx_converter = PDFToPPTXConverter(license_key)
    return _pptx_converter


def convert_pdf_to_pptx(pdf_path: str, output_path: Optional[str] = None, 
                       license_key: Optional[str] = None) -> Tuple[bool, str]:
    """便捷函数：PDF转PPTX"""
    converter = get_pptx_converter(license_key)
    return converter.convert_pdf_to_pptx(pdf_path, output_path)


def check_sdk_status(license_key: Optional[str] = None) -> dict:
    """检查SDK状态"""
    converter = get_pptx_converter(license_key)
    return converter.get_sdk_info()


if __name__ == "__main__":
    # 测试代码
    converter = PDFToPPTXConverter()
    print("SDK状态:", converter.get_sdk_info())
