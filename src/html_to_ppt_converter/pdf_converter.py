#!/usr/bin/env python3
"""
简化的HTML到PDF转换器
基于Pyppeteer，专门用于PPT幻灯片转换
"""

import os
import asyncio
import logging
from pathlib import Path
from typing import List, Optional, Dict, Any
import tempfile

try:
    from playwright.async_api import async_playwright
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False

# 保持向后兼容
PYPPETEER_AVAILABLE = PLAYWRIGHT_AVAILABLE

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class HTMLToPDFConverter:
    """HTML到PDF转换器"""

    def __init__(self):
        self.playwright = None
        self.browser = None

    def is_available(self) -> bool:
        """检查Playwright是否可用"""
        return PLAYWRIGHT_AVAILABLE

    async def start_browser(self):
        """启动浏览器"""
        if not self.is_available():
            raise RuntimeError("Playwright not available. Please install: pip install playwright")

        if self.browser is None:
            logger.info("启动Chromium浏览器...")
            self.playwright = await async_playwright().start()
            self.browser = await self.playwright.chromium.launch(
                headless=True,
                args=[
                    '--no-sandbox',
                    '--disable-setuid-sandbox',
                    '--disable-dev-shm-usage',
                    '--disable-gpu',
                    '--disable-web-security',
                    '--allow-running-insecure-content',
                    '--disable-features=VizDisplayCompositor'
                ]
            )

    async def close_browser(self):
        """关闭浏览器"""
        if self.browser:
            await self.browser.close()
            self.browser = None
        if self.playwright:
            await self.playwright.stop()
            self.playwright = None
    
    async def convert_html_to_pdf(self, html_file_path: str, pdf_output_path: str) -> bool:
        """
        将HTML文件转换为PDF
        
        Args:
            html_file_path: HTML文件路径
            pdf_output_path: PDF输出路径
            
        Returns:
            bool: 转换是否成功
        """
        if not os.path.exists(html_file_path):
            logger.error(f"HTML文件不存在: {html_file_path}")
            return False
        
        try:
            await self.start_browser()
            page = await self.browser.new_page()

            # 设置视口大小为PPT标准尺寸 (16:9)
            await page.set_viewport_size({'width': 1280, 'height': 720})

            # 加载HTML文件
            file_url = f"file://{os.path.abspath(html_file_path)}"
            logger.info(f"加载HTML文件: {file_url}")

            await page.goto(file_url, wait_until='networkidle', timeout=30000)

            # 等待页面完全加载
            await asyncio.sleep(2)

            # 执行页面就绪检查
            await self._wait_for_page_ready(page)

            # PDF生成选项 - PPT标准尺寸 (16:9)
            await page.pdf(
                path=pdf_output_path,
                width='338.67mm',  # 1280px at 96dpi
                height='190.5mm',  # 720px at 96dpi
                print_background=True,
                landscape=False,
                margin={
                    'top': '0mm',
                    'right': '0mm',
                    'bottom': '0mm',
                    'left': '0mm'
                },
                prefer_css_page_size=False,
                display_header_footer=False,
                scale=1
            )
            await page.close()

            logger.info(f"PDF生成成功: {pdf_output_path}")
            return True
            
        except Exception as e:
            logger.error(f"HTML到PDF转换失败: {e}")
            return False
    
    async def _wait_for_page_ready(self, page):
        """等待页面完全就绪"""
        try:
            # 等待所有图片加载完成
            await page.evaluate('''
                () => {
                    return new Promise((resolve) => {
                        const images = document.querySelectorAll('img');
                        let loadedCount = 0;
                        const totalImages = images.length;

                        if (totalImages === 0) {
                            resolve();
                            return;
                        }

                        images.forEach(img => {
                            if (img.complete) {
                                loadedCount++;
                            } else {
                                img.onload = img.onerror = () => {
                                    loadedCount++;
                                    if (loadedCount === totalImages) {
                                        resolve();
                                    }
                                };
                            }
                        });

                        if (loadedCount === totalImages) {
                            resolve();
                        }
                    });
                }
            ''')

            # 等待字体加载
            await page.evaluate('document.fonts.ready')

            # 额外等待确保渲染完成
            await asyncio.sleep(1)

        except Exception as e:
            logger.warning(f"页面就绪检查出现警告: {e}")
    
    async def convert_multiple_html_to_pdf(self, html_files: List[str], output_dir: str, 
                                         merged_pdf_path: Optional[str] = None) -> List[str]:
        """
        批量转换HTML文件为PDF
        
        Args:
            html_files: HTML文件路径列表
            output_dir: PDF输出目录
            merged_pdf_path: 合并后的PDF路径（可选）
            
        Returns:
            List[str]: 生成的PDF文件路径列表
        """
        os.makedirs(output_dir, exist_ok=True)
        pdf_files = []
        
        try:
            await self.start_browser()
            
            for i, html_file in enumerate(html_files):
                pdf_filename = f"slide_{i+1}.pdf"
                pdf_path = os.path.join(output_dir, pdf_filename)
                
                success = await self.convert_html_to_pdf(html_file, pdf_path)
                if success:
                    pdf_files.append(pdf_path)
                else:
                    logger.error(f"转换失败: {html_file}")
            
            # 如果需要合并PDF
            if merged_pdf_path and pdf_files:
                await self._merge_pdfs(pdf_files, merged_pdf_path)
                return [merged_pdf_path]
            
            return pdf_files
            
        except Exception as e:
            logger.error(f"批量转换失败: {e}")
            return []
        finally:
            await self.close_browser()
    
    async def _merge_pdfs(self, pdf_files: List[str], output_path: str):
        """合并多个PDF文件"""
        try:
            from PyPDF2 import PdfMerger
            
            merger = PdfMerger()
            for pdf_file in pdf_files:
                if os.path.exists(pdf_file):
                    merger.append(pdf_file)
            
            with open(output_path, 'wb') as output_file:
                merger.write(output_file)
            
            merger.close()
            logger.info(f"PDF合并成功: {output_path}")
            
        except ImportError:
            logger.warning("PyPDF2未安装，无法合并PDF。请安装: pip install PyPDF2")
        except Exception as e:
            logger.error(f"PDF合并失败: {e}")


# 全局转换器实例
_pdf_converter = None

def get_pdf_converter() -> HTMLToPDFConverter:
    """获取PDF转换器实例"""
    global _pdf_converter
    if _pdf_converter is None:
        _pdf_converter = HTMLToPDFConverter()
    return _pdf_converter


async def convert_html_to_pdf(html_file_path: str, pdf_output_path: str) -> bool:
    """便捷函数：单个HTML转PDF"""
    converter = get_pdf_converter()
    return await converter.convert_html_to_pdf(html_file_path, pdf_output_path)


async def convert_multiple_html_to_pdf(html_files: List[str], output_dir: str, 
                                     merged_pdf_path: Optional[str] = None) -> List[str]:
    """便捷函数：批量HTML转PDF"""
    converter = get_pdf_converter()
    return await converter.convert_multiple_html_to_pdf(html_files, output_dir, merged_pdf_path)
