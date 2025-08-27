#!/usr/bin/env python3
"""
HTML到PPT转换器配置文件
"""

import os
from pathlib import Path
from typing import Optional, Dict, Any
import json

class Config:
    """转换器配置类"""
    
    def __init__(self, config_file: Optional[str] = None):
        self.config_file = config_file or "converter_config.json"
        self.config_data = {}
        self.load_config()
    
    def load_config(self):
        """加载配置文件"""
        config_path = Path(self.config_file)
        
        if config_path.exists():
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    self.config_data = json.load(f)
                print(f"配置文件加载成功: {config_path}")
            except Exception as e:
                print(f"配置文件加载失败: {e}")
                self.config_data = {}
        else:
            print(f"配置文件不存在，将使用默认配置: {config_path}")
            self.create_default_config()
    
    def create_default_config(self):
        """创建默认配置文件"""
        default_config = {
            "apryse_sdk": {
                "license_key": "",
                "note": "请在此处填入您的Apryse SDK许可证密钥"
            },
            "pdf_options": {
                "width": "338.67mm",
                "height": "190.5mm",
                "print_background": True,
                "landscape": False,
                "margin": {
                    "top": "0mm",
                    "right": "0mm",
                    "bottom": "0mm",
                    "left": "0mm"
                }
            },
            "browser_options": {
                "headless": True,
                "timeout": 30000,
                "wait_for_images": True,
                "wait_for_fonts": True,
                "extra_wait_time": 2
            },
            "output": {
                "pdf_quality": "high",
                "merge_pdfs": True,
                "cleanup_temp_files": True
            },
            "slides": {
                "input_directory": "slides",
                "output_directory": "output",
                "index_file": "index.html"
            }
        }
        
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, indent=4, ensure_ascii=False)
            print(f"默认配置文件已创建: {self.config_file}")
            self.config_data = default_config
        except Exception as e:
            print(f"创建默认配置文件失败: {e}")
            self.config_data = default_config
    
    def get(self, key: str, default: Any = None) -> Any:
        """获取配置值（支持点号分隔的嵌套键）"""
        keys = key.split('.')
        value = self.config_data
        
        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return default
        
        return value
    
    def set(self, key: str, value: Any):
        """设置配置值（支持点号分隔的嵌套键）"""
        keys = key.split('.')
        config = self.config_data
        
        for k in keys[:-1]:
            if k not in config:
                config[k] = {}
            config = config[k]
        
        config[keys[-1]] = value
    
    def save_config(self):
        """保存配置到文件"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config_data, f, indent=4, ensure_ascii=False)
            print(f"配置已保存: {self.config_file}")
        except Exception as e:
            print(f"保存配置失败: {e}")
    
    @property
    def apryse_license_key(self) -> Optional[str]:
        """获取Apryse SDK许可证密钥"""
        key = self.get('apryse_sdk.license_key', '').strip()
        return key if key else None
    
    @property
    def slides_input_dir(self) -> str:
        """获取幻灯片输入目录"""
        return self.get('slides.input_directory', 'slides')
    
    @property
    def output_dir(self) -> str:
        """获取输出目录"""
        return self.get('slides.output_directory', 'output')
    
    @property
    def index_file(self) -> str:
        """获取索引文件名"""
        return self.get('slides.index_file', 'index.html')
    
    @property
    def pdf_options(self) -> Dict[str, Any]:
        """获取PDF生成选项"""
        return self.get('pdf_options', {})
    
    @property
    def browser_options(self) -> Dict[str, Any]:
        """获取浏览器选项"""
        return self.get('browser_options', {})
    
    @property
    def should_merge_pdfs(self) -> bool:
        """是否合并PDF"""
        return self.get('output.merge_pdfs', True)
    
    @property
    def should_cleanup_temp_files(self) -> bool:
        """是否清理临时文件"""
        return self.get('output.cleanup_temp_files', True)
    
    def validate_config(self) -> Dict[str, Any]:
        """验证配置"""
        issues = []
        warnings = []
        
        # 检查Apryse许可证密钥
        if not self.apryse_license_key:
            warnings.append("未设置Apryse SDK许可证密钥，将使用试用模式")
        
        # 检查输入目录
        slides_dir = Path(self.slides_input_dir)
        if not slides_dir.exists():
            issues.append(f"幻灯片输入目录不存在: {slides_dir}")
        
        # 检查索引文件
        index_path = slides_dir / self.index_file
        if slides_dir.exists() and not index_path.exists():
            warnings.append(f"索引文件不存在: {index_path}")
        
        return {
            "valid": len(issues) == 0,
            "issues": issues,
            "warnings": warnings
        }
    
    def print_config_summary(self):
        """打印配置摘要"""
        print("\n=== 转换器配置摘要 ===")
        print(f"输入目录: {self.slides_input_dir}")
        print(f"输出目录: {self.output_dir}")
        print(f"索引文件: {self.index_file}")
        print(f"Apryse许可证: {'已设置' if self.apryse_license_key else '未设置（试用模式）'}")
        print(f"合并PDF: {'是' if self.should_merge_pdfs else '否'}")
        print(f"清理临时文件: {'是' if self.should_cleanup_temp_files else '否'}")
        
        validation = self.validate_config()
        if validation['warnings']:
            print("\n警告:")
            for warning in validation['warnings']:
                print(f"  - {warning}")
        
        if validation['issues']:
            print("\n错误:")
            for issue in validation['issues']:
                print(f"  - {issue}")
        
        print("=" * 25)


# 全局配置实例
_config = None

def get_config(config_file: Optional[str] = None) -> Config:
    """获取配置实例"""
    global _config
    if _config is None:
        _config = Config(config_file)
    return _config


def load_config(config_file: Optional[str] = None) -> Config:
    """加载配置"""
    return get_config(config_file)


if __name__ == "__main__":
    # 测试配置
    config = Config()
    config.print_config_summary()
    
    # 验证配置
    validation = config.validate_config()
    print(f"\n配置验证结果: {'通过' if validation['valid'] else '失败'}")
