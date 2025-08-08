# WPS Parse  （.wps&.doc文件转DOCX和Markdown）

一个用于将 WPS（DOC） 文件转换为 DOCX 和 Markdown 格式的 Python 工具库。

.wps 和 .doc 都是“黑盒”二进制，所以这两种格式都支持。

## 功能特性

- 🔄 **WPS（DOC） 转 DOCX**：将 .wps 文件转换为 Microsoft Word .docx 格式
- 📝 **WPS（DOC） 转 Markdown**：将 .wps 文件转换为 Markdown .md 格式
- 🧹 **智能文本清理**：自动过滤乱码和不可读字符
- 📊 **可读性检测**：基于字符类型比例判断文本可读性
- 🛡️ **错误处理**：完善的异常处理机制

## 安装依赖

```bash
pip install -r requirements.txt
```

### 依赖库

- `olefile==0.47` - 用于解析 OLE 文件格式
- `docx==0.2.4` - 基础 DOCX 操作库
- `python-docx==1.2.0` - 高级 DOCX 文档处理库

## 使用方法

### 🖥️ GUI界面使用（推荐），📦 独立exe程序

将现有功能已经打包成exe文件，在./dist目录下，直接点击使用即可。

```bash
# 有python环境的盆友，可以选择直接运行
python gui_app.py

# 没有python环境的盆友，直接运行打包好的exe程序
./dist/WPS转换工具.exe
```

界面功能：

- 📁 可视化文件选择
- 🔄 支持批量转换
- 📊 实时进度显示
- ⚙️ 可调节可读性阈值
- 📝 详细转换日志

### 💻 命令行使用

#### WPS 转 DOCX

```python
from pathlib import Path
from wps_parse.wps_to_docx import wps_to_docx

# 设置输入和输出路径
src = Path('input.wps')
dst = Path('output.docx')

# 执行转换
wps_to_docx(src, dst)
```

#### WPS 转 Markdown

```python
from pathlib import Path
from wps_parse.wps_to_markdown import wps_to_md

# 设置输入和输出路径
src = Path('input.wps')
dst = Path('output.md')

# 执行转换
wps_to_md(src, dst)
```

#### 自定义可读性阈值

```python
# 调整可读性检测阈值（默认为 1.0，即 100% 可读字符）
wps_to_docx(src, dst, readability_threshold=0.8)  # 80% 可读字符
wps_to_md(src, dst, readability_threshold=0.8)
```

## 项目结构

```
Wps_Parse/
├── wps_parse/              # 主要模块
│   ├── __init__.py
│   ├── wps_to_docx.py      # WPS（DOC） 转 DOCX 功能
│   └── wps_to_markdown.py  # WPS（DOC） 转 Markdown 功能
├── dist/ 
│   └── WPS转换工具.exe      # 打包好的exe程序
├── gui_app.py              # 可视化GUI界面，可直接运行
│ 
├── requirements.txt        # 依赖列表
└── README.md              # 项目说明
```

## 核心功能说明

### 文本提取与清理

1. **OLE 文件解析**：使用 `olefile` 库解析 WPS（DOC） 文件的 OLE 容器结构
2. **WordDocument 流提取**：从 OLE 容器中提取 WordDocument 数据流
3. **编码处理**：使用 UTF-16LE 编码解码文本数据
4. **字符清理**：移除控制字符、NULL 字符、代理对和私用区字符
5. **换行统一**：将不同格式的换行符统一为 `\n`

### 可读性检测

通过分析文本中可打印字符的比例来判断文本是否为可读：

- 字母、数字、汉字、标点符号、空格等被视为可读字符
- 默认阈值为 100%，可根据需要调整
- 自动过滤掉乱码段落

### 输出格式

- **DOCX 格式**：每个段落作为独立的 Word 段落
- **Markdown 格式**：使用文件名作为一级标题，每个段落用双换行分隔

## 注意事项

1. **文件格式支持**：目前仅支持标准的 WPS & DOC 文件格式（OLE 容器）
2. **编码假设**：假设 WPS & DOC 文件使用 UTF-16LE 编码
3. **格式限制**：转换过程中会丢失原文档的格式信息（字体、颜色、样式等）
4. **文本内容**：主要提取纯文本内容，不包括图片、表格等复杂元素
