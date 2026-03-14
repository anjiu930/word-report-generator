# Word 报告编写助手 (word-report)

## 概述

Word 报告编写助手是一个用于自动生成格式化 Word 报告的 Claude Skill。该工具可以帮助用户快速生成课程设计报告、实验报告、项目报告等各种格式规范的 Word 文档。

## 前置条件

### 1. 软件依赖

| 软件 | 版本要求 | 说明 |
|------|----------|------|
| Python | 3.7+ | 报告生成脚本运行环境 |
| python-docx | 最新版 | Word 文档生成库 |

安装命令：
```bash
pip install python-docx
```

### 2. 文件依赖

| 文件 | 路径 | 说明 |
|------|------|------|
| template.py | `C:\Users\31930\.claude\skills\word-report\template.py` | 报告生成脚本模板 |
| 课程设计报告书模板.docx | `C:\Users\31930\.claude\skills\word-report\examples\` | Word 模板文件 |
| analyze_template.py | `C:\Users\31930\.claude\skills\word-report\` | 模板分析脚本 |

### 3. 目录要求

| 目录 | 路径 | 说明 |
|------|------|------|
| 脚本输出目录 | `D:\report\py\` | Python 脚本存放位置 |
| 报告输出目录 | `D:\report\word\` | 生成的 Word 报告存放位置 |

> 注意：这两个目录会在脚本运行时自动创建，无需手动创建。

## 环境配置

### 1. Python 环境配置

确保系统已安装 Python 3.7 或更高版本。可以通过以下命令检查：
```bash
python --version
```

### 2. 依赖库安装

在命令行中执行以下命令安装所需依赖：
```bash
pip install python-docx
```

### 3. Claude  code配置

确保 Claude code 已正确安装并配置了以下 MCP 工具：
- chrome-devtools（用于知网文献搜索）
- desktop-commander（用于文件操作）（可选）

## 使用流程

### 步骤 1：用户请求

用户需要提供以下信息：
1. 项目路径：包含源码的文件夹路径
2. 参考文档模板（可选）：如果用户有特定的报告模板，提供其路径

### 步骤 2：模板分析

使用 analyze_template.py 分析用户提供的模板，提取章节标题结构：
```bash
python analyze_template.py "模板路径"
```

### 步骤 3：文献搜索

使用 Chrome DevTools MCP 访问知网搜索相关文献：
1. 访问 https://www.cnki.net
2. 输入简短关键词搜索（如"嵌入式系统 传感器"）
3. 筛选 8 篇近 5 年的相关文献

### 步骤 4：编写脚本

根据模板结构，编写报告生成脚本。脚本必须：
- 参考 template.py 的代码结构
- 只修改中文内容，不更改代码结构
- 使用提供的常用函数

### 步骤 5：生成报告

运行脚本生成 Word 报告：
```bash
python D:\report\py\报告脚本.py
```

生成的报告保存在 `D:\report\word\` 目录下。

## 常用函数

| 函数 | 用途 |
|------|------|
| `add_chapter_title(doc, "第X章 标题")` | 添加一级标题（章标题） |
| `add_section_title(doc, "X.X 标题")` | 添加二级标题（小节标题） |
| `add_body_paragraph(doc, "正文内容")` | 添加正文段落（首行缩进2字符） |
| `add_image_placeholder(doc, "图片描述", 编号)` | 添加图片占位符 |
| `add_code_with_description(doc, "功能描述", "代码")` | 添加带描述的代码块 |
| `add_reference_item(doc, "参考文献内容")` | 添加参考文献条目 |

## 格式规范

### 标题层级
- 一级标题（章标题）：字号 15pt，加粗
- 二级标题（小节标题）：字号 12pt，加粗

### 正文段落
- 字体：宋体
- 字号：五号（10.5pt）
- 行距：1.25 倍行距
- 首行缩进：2 字符

### 图片规范
- 放置位置：仅限设计与实现章节
- 引用格式：如"如图X所示"
- 命名格式：`图X 图片描述`

### 代码块
- 字体：宋体
- 字号：9pt
- 左侧对齐，无首行缩进

## 重要规则

1. **模板路径禁止修改**：绝对不能修改 template.py 中的 TEMPLATE_PATH
2. **禁止使用 Markdown**：不要在正文中使用 ###、**、--、# 等标记
3. **禁止三级标题**：只使用一级和二级标题
4. **图片引用**：每张图片必须在上一段文字中有引用说明

## 故障排除

### 问题 1：报告生成失败

**可能原因**：python-docx 库未安装
**解决方法**：执行 `pip install python-docx`

### 问题 2：模板路径错误

**可能原因**：TEMPLATE_PATH 指向的模板文件不存在
**解决方法**：检查 `C:\Users\31930\.claude\skills\word-report\examples\` 目录

### 问题 3：输出目录无权限

**可能原因**：D:\report\ 目录无写入权限
**解决方法**：以管理员身份运行或更换输出目录

## 扩展使用

### 自定义模板

如果需要使用自定义模板：
  给出参考模版啊地址即可
### 添加新功能

如需扩展报告功能，可以：
1. 在 template.py 中添加新的辅助函数
2. 确保遵循现有的代码结构风格
3. 更新本 README 文档


