# HTML到PPT转换器

一个独立的HTML幻灯片到PowerPoint转换工具，支持将HTML格式的演示文稿转换为标准的PPTX文件。

## 功能特点

- 🔄 **两步转换**: HTML → PDF → PPTX
- 📊 **完美保持样式**: 支持CSS样式、图表、动画等
- 📏 **标准PPT尺寸**: 16:9 (1280x720) 标准演示尺寸
- 🔧 **灵活配置**: 支持自定义配置文件
- 📁 **批量处理**: 自动扫描并转换多个幻灯片
- 🧹 **自动清理**: 可选的临时文件清理

## 安装依赖

### 基础依赖
```bash
# 使用uv安装Python依赖
uv add pyppeteer PyPDF2

# 或使用pip
pip install pyppeteer PyPDF2
```

### Apryse SDK
1. 下载Apryse SDK (原PDFTron SDK)
2. 按照官方文档安装
3. 获取许可证密钥（可选，无密钥将使用试用模式）

## 目录结构

```
src/html_to_ppt_converter/
├── main.py                 # 主程序
├── config.py              # 配置管理
├── pdf_converter.py       # HTML到PDF转换器
├── pptx_converter.py      # PDF到PPTX转换器
├── converter_config.json  # 配置文件（自动生成）
├── README.md              # 说明文档
└── slides/                # HTML幻灯片目录（需要创建）
    ├── index.html         # 索引页面（可选）
    ├── slide1.html        # 第1张幻灯片
    ├── slide2.html        # 第2张幻灯片
    └── ...
```

## 使用方法

### 1. 准备HTML文件
将您的HTML幻灯片文件放入 `slides/` 目录：
- `index.html` - 索引页面（可选）
- `slide1.html`, `slide2.html`, ... - 各个幻灯片

### 2. 配置转换器
首次运行会自动生成 `converter_config.json` 配置文件：

```json
{
    "apryse_sdk": {
        "license_key": "在此填入您的Apryse SDK许可证密钥",
        "note": "请在此处填入您的Apryse SDK许可证密钥"
    },
    "pdf_options": {
        "width": "338.67mm",
        "height": "190.5mm",
        "print_background": true,
        "landscape": false,
        "margin": {
            "top": "0mm",
            "right": "0mm",
            "bottom": "0mm",
            "left": "0mm"
        }
    },
    "slides": {
        "input_directory": "slides",
        "output_directory": "output",
        "index_file": "index.html"
    }
}
```

### 3. 运行转换器
```bash
# 使用uv运行
cd src/html_to_ppt_converter
uv run python main.py

# 或直接运行
python main.py
```

### 4. 查看结果
转换完成后，PPTX文件将保存在 `output/` 目录中。

## 配置选项

### Apryse SDK配置
- `license_key`: Apryse SDK许可证密钥（可选）

### PDF生成选项
- `width/height`: PDF页面尺寸（默认16:9）
- `print_background`: 是否包含背景
- `margin`: 页面边距

### 浏览器选项
- `headless`: 无头模式（默认true）
- `timeout`: 页面加载超时时间
- `wait_for_images`: 等待图片加载
- `wait_for_fonts`: 等待字体加载

### 输出选项
- `merge_pdfs`: 是否合并多个PDF
- `cleanup_temp_files`: 是否清理临时文件

## 转换流程

1. **扫描HTML文件**: 自动查找slides目录中的HTML文件
2. **HTML转PDF**: 使用Pyppeteer将HTML渲染为PDF
3. **PDF转PPTX**: 使用Apryse SDK将PDF转换为PowerPoint
4. **清理临时文件**: 可选的临时文件清理

## 故障排除

### 常见问题

1. **Pyppeteer安装失败**
   ```bash
   # 手动安装Chromium
   pyppeteer-install
   ```

2. **Apryse SDK不可用**
   - 检查SDK是否正确安装
   - 验证许可证密钥是否有效
   - 无许可证将使用试用模式（有水印）

3. **HTML文件未找到**
   - 确保slides目录存在
   - 检查文件命名是否正确
   - 查看日志文件了解详细信息

4. **转换质量问题**
   - 调整PDF生成选项
   - 检查HTML中的CSS样式
   - 确保所有资源文件可访问

### 日志文件
程序运行时会生成 `converter.log` 日志文件，包含详细的执行信息。

## 技术细节

- **PDF转换器**: 基于Pyppeteer (Chromium)
- **PPTX转换器**: 基于Apryse SDK
- **支持格式**: HTML, CSS, JavaScript, 图片, 图表
- **输出格式**: 标准PPTX文件，可在PowerPoint中编辑

## 许可证

本工具使用的第三方库：
- Pyppeteer: MIT License
- Apryse SDK: 商业许可证（需要购买）
- PyPDF2: BSD License
