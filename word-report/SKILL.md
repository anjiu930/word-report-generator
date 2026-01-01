---
name: word-report
description: Word 报告编写助手。当用户说"编写报告"、"帮我写一份报告"、"生成报告"、"创建 Word 文档"等关键词时调用此 Skill。使用 python-docx 库生成格式化的 Word 报告文档。
allowed-tools: Bash, Write, Read, Task, AskUserQuestion, TodoWrite
---

# ⚠️ 执行前必读

## 🚨 核心原则（必须遵守）

**参考 template.py 编写报告，只对中文内容进行修改，不允许更改已有代码段！**

- ✅ **允许修改**：中文文字内容（项目描述、段落文字等）
- ❌ **禁止修改**：函数定义、代码结构、格式设置、缩进等

**违反此原则将导致报告格式错误！**

---

**本 Skill 依赖于以下配置文件，执行前必须先读取：**
任务优先级:skill.md=guide.md->examples

| 文件/文件夹 | 路径 | 用途 |
|-------------|------|------|
| **guide.md** | `C:\Users\31930\.claude\skills\word-report\guide.md` | 详细的 A4 报告编写要求（核心必读） |
| **template.py** | `C:\Users\31930\.claude\skills\word-report\template.py` | Python 脚本模板（完全采取现有脚本格式，只修改中文内容） |
| **examples/** | `C:\Users\31930\.claude\skills\word-report\examples\` | 参考示例文档与官方报告任务书，报告模板 |


**执行步骤：**
1. 首先读取 `guide.md` 熟悉格式规范
2. 参阅 `template.py` 了解代码模板
3. 查看 `examples/` 中的示例文档作为参考
4. 使用 Chrome DevTools MCP 搜索与项目核心内容相关的参考文献（国内文献，8篇，近5年内高质量文献）
5. 按照规范编写报告脚本，严格确保word报告整体字符在6800字左右，生成的脚本内容放到此路径'E:\report\py'
6. 执行脚本生成文档,生成的报告内容放到此路径'E:\report\word'

---

# Word 报告编写助手

## 功能说明
这是一个帮助用户快速生成格式化 Word 报告的 Skill。支持创建包含标题、正文、表格、图片等元素的报告文档。

## 工作流程

1. **确认需求**
   - 根据用户提供的项目目录，或者询问用户报告的主题和内容要求

2. **根据 Python 脚本来进行编写**
   - 保持原有脚本格式不变，只在脚本中填充或更换中文文字内容来进行编写
   - 严格确保word报告整体字符在6800字左右
   
3. **执行脚本**
   - 运行 Python 脚本生成 Word 文档
   - 确认文件创建成功

4. **结果反馈**
   - 告知用户报告已生成
   - 提供文件路径

## 常用代码模板

### 基础报告结构

```python
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 创建文档
doc = Document()

# 添加正文段落
doc.add_paragraph('这里是报告内容...')

# 保存文档
doc.save('报告.docx')
```
### 添加图片

```python
from docx.shared import Inches

# 添加图片
doc.add_picture('图片路径.png', width=Inches(4.0))
```

## 交互模板

当触发此 Skill 时，使用以下交互流程：

```
用户: "帮我写一份报告"
你: "好的宝贝～让我帮你生成 Word 报告！

💡 提示：你可以在 examples/ 目录下的课程设计报告模板中预先填写个人与课程信息，
这样可以避免每次生成报告时手动修改封面。

请告诉我：
1. 报告的主题或者项目的要求是什么？
2. 需要包含哪些内容部分？
```

## 注意事项

- 文件名使用中文时确保编码正确
- 生成的文档使用 `.docx` 格式

---

## 📁 文件输出位置

- **脚本路径**：`E:\report\py\` - 生成的 Python 脚本存放位置
- **报告路径**：`E:\report\word\` - 生成的 Word 报告存放位置
- **文件格式**：`.docx`

---

## 🔍 参考文献搜索

使用 Chrome DevTools MCP 访问知网搜索文献，选择 8 篇近 5 年内与项目较相关的文献写入报告。

**操作流程：**
1. navigate_page - 访问 https://www.cnki.net
2. fill - 输入关键词
3. click - 点击检索按钮
4. take_snapshot - 获取搜索结果

**文献格式规范：**
```
[序号] 作者. 标题[J]. 期刊名称, 年份, 卷(期): 页码.
[序号] 作者. 标题[M]. 出版社: 年份.
[序号] 作者. 标题[EB/OL]. URL, 年份.
```


