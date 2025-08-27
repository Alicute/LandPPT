#!/usr/bin/env python3
"""
HTML到PPT转换器主程序
自动扫描slides目录并转换为PPT
"""

import os
import sys
import asyncio
import logging
from pathlib import Path
from typing import List, Optional
import tempfile
import shutil

# 添加当前目录到Python路径
sys.path.insert(0, os.path.dirname(__file__))

from config import get_config
from pdf_converter import get_pdf_converter
from pptx_converter import get_pptx_converter

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('converter.log', encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)


class HTMLToPPTConverter:
    """HTML到PPT转换器主类"""
    
    def __init__(self, config_file: Optional[str] = None):
        self.config = get_config(config_file)
        self.pdf_converter = get_pdf_converter()
        self.pptx_converter = get_pptx_converter(self.config.apryse_license_key)
    
    def find_html_files(self) -> List[str]:
        """查找slides目录中的HTML文件"""
        slides_dir = Path(self.config.slides_input_dir)

        if not slides_dir.exists():
            logger.error(f"slides目录不存在: {slides_dir}")
            return []

        # 优先查找slide_1.html, slide2.html等（实际幻灯片内容）
        slide_files = []
        for i in range(1, 1000):  # 最多支持999页
            slide_file = slides_dir / f"slide_{i}.html"
            if slide_file.exists():
                slide_files.append(str(slide_file))
            else:
                break

        if slide_files:
            logger.info(f"找到{len(slide_files)}个幻灯片文件 (slide_1~slide{len(slide_files)})")
            return slide_files

        # 如果没有找到slide1~N格式，查找index.html
        index_file = slides_dir / self.config.index_file
        if index_file.exists():
            logger.info(f"找到索引文件: {index_file}")
            return [str(index_file)]

        # 最后查找所有HTML文件（排除index.html）
        all_html = list(slides_dir.glob("*.html"))
        html_files = [str(f) for f in sorted(all_html) if f.name != self.config.index_file]

        if html_files:
            logger.info(f"找到{len(html_files)}个HTML文件")
        else:
            logger.warning("未找到任何HTML文件")

        return html_files
    
    async def convert_to_pdf(self, html_files: List[str]) -> Optional[str]:
        """将HTML文件转换为PDF"""
        if not html_files:
            logger.error("没有找到HTML文件")
            return None
        
        # 创建输出目录
        output_dir = Path(self.config.output_dir)
        output_dir.mkdir(exist_ok=True)
        
        # 创建临时PDF目录
        temp_pdf_dir = output_dir / "temp_pdfs"
        temp_pdf_dir.mkdir(exist_ok=True)
        
        try:
            # 生成带时间戳的文件名
            import datetime
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

            if self.config.should_merge_pdfs and len(html_files) > 1:
                # 批量转换并合并
                merged_pdf_path = str(output_dir / f"presentation_{timestamp}.pdf")
                pdf_files = await self.pdf_converter.convert_multiple_html_to_pdf(
                    html_files, str(temp_pdf_dir), merged_pdf_path
                )
                
                if pdf_files and os.path.exists(merged_pdf_path):
                    logger.info(f"PDF合并完成: {merged_pdf_path}")
                    return merged_pdf_path
                else:
                    logger.error("PDF转换或合并失败")
                    return None
            else:
                # 单个文件转换
                pdf_path = str(output_dir / f"presentation_{timestamp}.pdf")
                success = await self.pdf_converter.convert_html_to_pdf(html_files[0], pdf_path)
                
                if success:
                    logger.info(f"PDF转换完成: {pdf_path}")
                    return pdf_path
                else:
                    logger.error("PDF转换失败")
                    return None
        
        finally:
            # 清理临时文件
            if self.config.should_cleanup_temp_files and temp_pdf_dir.exists():
                try:
                    shutil.rmtree(temp_pdf_dir)
                    logger.info("临时PDF文件已清理")
                except Exception as e:
                    logger.warning(f"清理临时文件失败: {e}")
    
    def convert_to_pptx(self, pdf_path: str) -> Optional[str]:
        """将PDF转换为PPTX"""
        if not os.path.exists(pdf_path):
            logger.error(f"PDF文件不存在: {pdf_path}")
            return None
        
        # 生成PPTX输出路径 - 使用PDF文件名对应的时间戳
        pdf_filename = Path(pdf_path).stem  # 获取不带扩展名的文件名
        output_dir = Path(self.config.output_dir)
        pptx_path = str(output_dir / f"{pdf_filename}.pptx")
        
        # 执行转换
        success, result = self.pptx_converter.convert_pdf_to_pptx(pdf_path, pptx_path)
        
        if success:
            logger.info(f"PPTX转换完成: {result}")
            return result
        else:
            logger.error(f"PPTX转换失败: {result}")
            return None
    
    async def convert(self) -> bool:
        """执行完整的转换流程"""
        logger.info("开始HTML到PPT转换流程")
        
        # 验证配置
        validation = self.config.validate_config()
        if not validation['valid']:
            logger.error("配置验证失败:")
            for issue in validation['issues']:
                logger.error(f"  - {issue}")
            return False
        
        if validation['warnings']:
            for warning in validation['warnings']:
                logger.warning(warning)
        
        # 检查转换器可用性
        if not self.pdf_converter.is_available():
            logger.error("PDF转换器不可用，请安装pyppeteer: pip install pyppeteer")
            return False
        
        if not self.pptx_converter.is_available():
            logger.error("PPTX转换器不可用，请安装Apryse SDK")
            return False
        
        # 查找HTML文件
        html_files = self.find_html_files()
        if not html_files:
            logger.error("未找到HTML文件")
            return False
        
        logger.info(f"找到{len(html_files)}个HTML文件，开始转换...")
        
        # 步骤1: HTML转PDF
        logger.info("步骤1: 转换HTML为PDF")
        pdf_path = await self.convert_to_pdf(html_files)
        if not pdf_path:
            logger.error("HTML到PDF转换失败")
            return False
        
        # 步骤2: PDF转PPTX
        logger.info("步骤2: 转换PDF为PPTX")
        pptx_path = self.convert_to_pptx(pdf_path)
        if not pptx_path:
            logger.error("PDF到PPTX转换失败")
            return False
        
        # 清理中间PDF文件（如果需要）
        if self.config.should_cleanup_temp_files:
            try:
                os.unlink(pdf_path)
                logger.info("中间PDF文件已清理")
            except Exception as e:
                logger.warning(f"清理中间文件失败: {e}")
        
        logger.info(f"转换完成！输出文件: {pptx_path}")
        return True
    
    def print_status(self):
        """打印转换器状态"""
        print("\n=== HTML到PPT转换器状态 ===")
        
        # 配置状态
        self.config.print_config_summary()
        
        # 转换器状态
        print(f"\nPDF转换器: {'可用' if self.pdf_converter.is_available() else '不可用'}")
        
        pptx_info = self.pptx_converter.get_sdk_info()
        print(f"PPTX转换器: {'可用' if pptx_info['available'] else '不可用'}")
        if pptx_info['available']:
            print(f"  - 版本: {pptx_info.get('version', '未知')}")
            print(f"  - 许可证: {'已提供' if pptx_info['license_provided'] else '未提供（试用模式）'}")
        
        # 文件状态
        html_files = self.find_html_files()
        print(f"\n找到HTML文件: {len(html_files)}个")
        for i, file in enumerate(html_files[:5]):  # 只显示前5个
            print(f"  {i+1}. {Path(file).name}")
        if len(html_files) > 5:
            print(f"  ... 还有{len(html_files)-5}个文件")
        
        print("=" * 30)


async def main():
    """主函数"""
    print("HTML到PPT转换器")
    print("=" * 30)
    
    # 创建转换器实例
    converter = HTMLToPPTConverter()
    
    # 显示状态
    converter.print_status()
    
    # 询问是否继续
    try:
        response = input("\n是否开始转换？(y/N): ").strip().lower()
        if response not in ['y', 'yes', '是']:
            print("转换已取消")
            return
    except KeyboardInterrupt:
        print("\n转换已取消")
        return
    
    # 执行转换
    success = await converter.convert()
    
    if success:
        print("\n✅ 转换成功完成！")
    else:
        print("\n❌ 转换失败，请检查日志")
        sys.exit(1)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n程序被用户中断")
    except Exception as e:
        logger.error(f"程序执行出错: {e}")
        sys.exit(1)
