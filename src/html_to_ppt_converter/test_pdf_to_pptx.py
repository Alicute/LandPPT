#!/usr/bin/env python3
"""
测试PDF到PPTX转换功能
直接使用现有的PDF文件进行转换测试
"""

import os
import sys
from pathlib import Path

# 添加当前目录到Python路径
sys.path.insert(0, os.path.dirname(__file__))

from config import get_config
from pptx_converter import get_pptx_converter


def find_pdf_files():
    """查找可用的PDF文件"""
    pdf_files = []
    
    # 查找output目录中的PDF文件
    output_dir = Path("output")
    if output_dir.exists():
        pdf_files.extend(list(output_dir.glob("*.pdf")))
    
    # 查找当前目录中的PDF文件
    current_dir = Path(".")
    pdf_files.extend(list(current_dir.glob("*.pdf")))
    
    return pdf_files


def test_apryse_installation():
    """测试Apryse SDK安装"""
    print("=== 测试Apryse SDK安装 ===")

    try:
        import apryse_sdk
        print("✅ apryse_sdk 模块导入成功")

        try:
            from apryse_sdk.PDFNetPython import PDFNet
            print("✅ PDFNet 导入成功")

            # 获取配置中的许可证密钥
            try:
                config = get_config()
                license_key = config.apryse_license_key

                if license_key:
                    print(f"使用配置中的许可证密钥: {license_key[:20]}...")
                    PDFNet.Initialize(license_key)
                    print("✅ PDFNet 初始化成功（使用许可证密钥）")
                else:
                    print("配置中没有许可证密钥，使用试用模式")
                    PDFNet.Initialize()
                    print("✅ PDFNet 初始化成功（试用模式）")

                return True

            except Exception as e:
                print(f"❌ PDFNet 初始化失败: {e}")
                return False

        except ImportError as e:
            print(f"❌ PDFNet 导入失败: {e}")
            return False

    except ImportError as e:
        print(f"❌ apryse_sdk 模块导入失败: {e}")
        print("请安装 Apryse SDK:")
        print("1. 访问 https://www.pdftron.com/")
        print("2. 下载并安装 Apryse SDK")
        print("3. 或尝试: pip install apryse-sdk")
        return False


def test_pdf_to_pptx_conversion(pdf_path: str):
    """测试PDF到PPTX转换"""
    print(f"\n=== 测试PDF到PPTX转换 ===")
    print(f"输入PDF: {pdf_path}")
    
    # 生成输出路径
    pdf_file = Path(pdf_path)
    pptx_path = pdf_file.parent / f"{pdf_file.stem}_converted.pptx"
    
    print(f"输出PPTX: {pptx_path}")
    
    # 获取配置和转换器
    config = get_config()
    converter = get_pptx_converter(config.apryse_license_key)
    
    # 检查转换器状态
    print(f"转换器可用性: {'✅' if converter.is_available() else '❌'}")
    
    if not converter.is_available():
        print("转换器不可用，无法继续测试")
        return False
    
    # 执行转换
    print("开始转换...")
    success, result = converter.convert_pdf_to_pptx(pdf_path, str(pptx_path))
    
    if success:
        print(f"✅ 转换成功！")
        print(f"输出文件: {result}")
        
        # 检查文件大小
        if os.path.exists(result):
            file_size = os.path.getsize(result)
            print(f"文件大小: {file_size:,} 字节")
        
        return True
    else:
        print(f"❌ 转换失败: {result}")
        return False


def main():
    """主函数"""
    print("PDF到PPTX转换测试工具")
    print("=" * 40)
    
    # 1. 测试Apryse SDK安装
    if not test_apryse_installation():
        print("\n请先解决Apryse SDK安装问题")
        return
    
    # 2. 查找PDF文件
    print(f"\n=== 查找PDF文件 ===")
    pdf_files = find_pdf_files()
    
    if not pdf_files:
        print("❌ 未找到PDF文件")
        print("请确保以下位置有PDF文件:")
        print("- output/presentation.pdf")
        print("- 当前目录中的任何.pdf文件")
        return
    
    print(f"找到 {len(pdf_files)} 个PDF文件:")
    for i, pdf_file in enumerate(pdf_files, 1):
        print(f"  {i}. {pdf_file}")
    
    # 3. 选择要转换的PDF文件
    if len(pdf_files) == 1:
        selected_pdf = pdf_files[0]
        print(f"\n自动选择: {selected_pdf}")
    else:
        try:
            choice = input(f"\n请选择要转换的PDF文件 (1-{len(pdf_files)}): ").strip()
            if not choice:
                selected_pdf = pdf_files[0]
                print(f"使用默认选择: {selected_pdf}")
            else:
                index = int(choice) - 1
                if 0 <= index < len(pdf_files):
                    selected_pdf = pdf_files[index]
                else:
                    print("无效选择，使用第一个文件")
                    selected_pdf = pdf_files[0]
        except (ValueError, KeyboardInterrupt):
            print("使用第一个文件")
            selected_pdf = pdf_files[0]
    
    # 4. 执行转换测试
    success = test_pdf_to_pptx_conversion(str(selected_pdf))
    
    if success:
        print(f"\n🎉 测试完成！转换成功！")
    else:
        print(f"\n❌ 测试失败")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n测试被用户中断")
    except Exception as e:
        print(f"\n测试过程中出现错误: {e}")
        import traceback
        traceback.print_exc()
