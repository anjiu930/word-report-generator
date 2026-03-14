<div align="center">

# Word Report Generator

### 别再被课程设计报告的排版折磨了。

<!-- ![demo](docs/images/demo-preview.png) -->

一个 Claude Code Skill 插件 —— 一句话生成格式完美的 Word 课程设计报告。

**告别手动调字体、改行距、对缩进的噩梦。**

[![Release](https://img.shields.io/github/v/release/anjiu930/word-report-generator?style=flat-square)](https://github.com/anjiu930/word-report-generator/releases)
[![License](https://img.shields.io/badge/license-MIT-blue?style=flat-square)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.7+-green?style=flat-square)](https://www.python.org/)

</div>

---

## 这个项目解决什么问题？

如果你是理工科大学生，你一定经历过这些：

> 代码写了一周，报告排版改了三天。
>
> 宋体小四还是五号？行距 1.25 还是 1.5？首行缩进到底多少厘米？
>
> 好不容易排好了，老师说"参考文献格式不对，封面信息没改，图片编号对不上"……
>
> 最崩溃的是——下周还有一门课设报告要交。

**这个工具就是为了终结这些痛苦而生的。**

你只需要对 Claude Code 说一句话：

```
帮我写一份课程设计报告，项目在 D:\MyProject
```

然后它会自动：

1. 读取你的项目源码，理解你写了什么
2. 按照学术规范生成完整的 Word 报告
3. 自动搜索知网文献，填充参考文献
4. 所有格式一步到位——**你不需要打开 Word 改任何东西**

## 用之前 vs 用之后

| | 以前手写报告 | 用这个工具 |
|---|---|---|
| 排版格式 | 每次手动调字体、行距、缩进 | 全自动，严格符合规范 |
| 写作时间 | 8~15 小时 | 5 分钟生成初稿 |
| 参考文献 | 手动搜知网、手动排格式 | 自动搜索 + 自动格式化 |
| 代码展示 | 截图粘贴 / 手动排版代码 | 自动从源码提取 + 排版 |
| 图片引用 | 经常忘记标注"如图X所示" | 自动生成引用 + 占位符 |
| 跨课程复用 | 每次从零开始 | 换个模板就能复用 |

## 效果预览

<!-- 截图占位，稍后补充 -->

| 封面 | 正文 + 代码 | 参考文献 |
|:----:|:----------:|:--------:|
| ![封面](docs/images/cover.png) | ![正文](docs/images/body.png) | ![文献](docs/images/references.png) |

> 以上均为工具实际生成的报告截图。生成后你只需要替换图片占位符为真实截图即可提交。

## 它会帮你生成什么？

一份完整的课程设计报告，包含：

```
封面（学校、姓名、学号等）
    ↓
目录（自动保留模板目录）
    ↓
第1章 项目背景 ──────── 自动撰写
第2章 开发环境 ──────── 自动撰写
第3章 开发工具 ──────── 自动撰写
第4章 完成时间 ──────── 自动撰写
第5章 设计思想 ──────── 自动撰写
第6章 设计过程 ──────── 读取你的源码 + 硬件描述 + 代码展示
第7章 测试运行 ──────── 自动撰写 + 图片占位
第8章 评价与修订 ────── 自动撰写
第9章 设计体会 ──────── 自动撰写
    ↓
参考文献（8篇，自动从知网搜索）
    ↓
评分表（自动从模板保留）
```

**默认生成约 7000 字**，符合大多数课程设计报告的字数要求。

---

## 快速开始

### 你需要准备什么

- [Claude Code](https://docs.anthropic.com/en/docs/claude-code)（Anthropic 的 AI 命令行工具）
- Python 3.7+
- 安装 python-docx：`pip install python-docx`

### 三步安装

**1. 下载本项目**

```bash
git clone https://github.com/anjiu930/word-report-generator.git
```

或者直接去 [Releases](https://github.com/anjiu930/word-report-generator/releases) 下载 zip。

**2. 复制到 Claude Code Skills 目录**

Windows:
```bash
xcopy /E /I word-report "%USERPROFILE%\.claude\skills\word-report"
```

macOS / Linux:
```bash
cp -r word-report ~/.claude/skills/word-report
```

**3. 修改路径为你自己的**

打开 `~/.claude/skills/word-report/template.py`，把前两行改成你的路径：

```python
OUTPUT_PATH = r"你想要的报告输出路径\项目报告.docx"
TEMPLATE_PATH = r"你的用户目录\.claude\skills\word-report\examples\课程设计报告书模板.docx"
```

同样，打开 `SKILL.md`，把里面出现的路径也替换成你自己的。

**搞定！** 现在打开 Claude Code，说"帮我写报告"就可以了。

> 首次使用时 Claude Code 会提示授权 Skill，选择允许即可。也可以手动在 `~/.claude/settings.local.json` 的 `allowedTools` 中添加 `"Skill(word-report)"`。

---

## 使用方法

### 最简单的用法

```
帮我写一份课程设计报告，项目路径在 D:\MyProject
```

### 有老师给的模板？

```
帮我写报告，参考模板在 D:\模板\报告模板.docx，项目在 D:\MyProject
```

工具会自动提取模板中的章节标题，严格按照老师的模板结构来写。

### 不想每次填封面信息？

打开 `examples/课程设计报告书模板.docx`，提前把你的学校、姓名、学号等封面信息填好。以后每次生成报告，封面就自动带上了。

### 完整流程

```
你："帮我写报告"
    ↓
Claude：询问项目路径、是否有参考模板
    ↓
Claude：分析模板 → 展示章节大纲 → 等你确认
    ↓
你："确认，开始写"
    ↓
Claude：读源码 → 搜文献 → 写脚本 → 生成 Word
    ↓
报告生成完毕！
```

---

## 格式规范一览

生成的报告严格遵循以下学术排版标准：

| 元素 | 字体 | 字号 | 其他 |
|------|------|------|------|
| 一级标题 | 宋体 加粗 | 15pt | 段后 4pt |
| 二级标题 | 宋体 加粗 | 12pt | 段前 2pt |
| 正文段落 | 宋体 | 10.5pt（五号） | 首行缩进 2 字符，行距 1.25 |
| 代码块 | 宋体 | 9pt | 左对齐，行距 1.0 |
| 图片标注 | 宋体 | 10.5pt | 居中，格式"图X 描述" |

---

## 项目结构

```
word-report-generator/
├── README.md                         # 你正在看的这个文件
├── docs/images/                      # 截图存放目录
└── word-report/                      # Skill 核心目录
    ├── SKILL.md                      # Skill 定义（触发词 + 执行规则）
    ├── template.py                   # 报告模板脚本（14 个格式函数）
    ├── analyze_template.py           # 模板标题提取工具
    └── examples/
        └── 课程设计报告书模板.docx     # 默认模板
```

## 函数参考

`template.py` 内置了 14 个函数，Claude 会自动调用它们来组装报告：

<details>
<summary><b>点击展开完整函数列表</b></summary>

| 函数 | 说明 |
|------|------|
| `add_chapter_title(doc, text)` | 添加一级标题（如"第1章 需求分析"） |
| `add_section_title(doc, text)` | 添加二级标题（如"1.1 功能需求"） |
| `add_body_paragraph(doc, text)` | 添加正文段落（首行缩进 2 字符） |
| `add_summary_paragraph(doc, text)` | 添加总领段落（无缩进） |
| `add_image_placeholder(doc, desc, num)` | 添加图片占位符 |
| `add_image_placeholder_with_ref(doc, desc, num)` | 添加图片占位符 + 自动引用文字 |
| `add_code_block(doc, code)` | 添加代码块 |
| `add_code_with_description(doc, desc, code)` | 添加带描述的代码块 |
| `add_references_title(doc)` | 添加"参考文献"标题 |
| `add_reference_item(doc, text)` | 添加参考文献条目 |
| `find_toc_end(doc)` | 自动检测目录结束位置 |
| `remove_content_after_toc(doc, index)` | 清除目录后的模板内容 |
| `remove_table_by_index(doc, index)` | 删除指定表格 |
| `move_table_to_end(doc, index)` | 将评分表移到文档末尾 |

</details>

---

## 常见问题

<details>
<summary><b>报错 "No module named 'docx'"</b></summary>

```bash
pip install python-docx
```

注意：包名是 `python-docx`，不是 `docx`，别装错了。
</details>

<details>
<summary><b>模板路径找不到</b></summary>

检查 `template.py` 里的 `TEMPLATE_PATH` 是否正确指向 `examples/课程设计报告书模板.docx`。如果你移动了文件，路径也要跟着改。
</details>

<details>
<summary><b>想用自己学校的模板</b></summary>

把你学校的 .docx 模板放到 `examples/` 目录下，使用时告诉 Claude 模板路径就行。Skill 会自动分析其中的章节结构。
</details>

<details>
<summary><b>支持 macOS / Linux 吗？</b></summary>

支持。把 `SKILL.md` 和 `template.py` 中的 Windows 路径改成你系统的路径就行：

```python
TEMPLATE_PATH = "/Users/yourname/.claude/skills/word-report/examples/课程设计报告书模板.docx"
```
</details>

<details>
<summary><b>可以不搜知网文献吗？</b></summary>

可以。如果你没装 Chrome DevTools MCP，Claude 会跳过文献搜索。你也可以手动给它一份文献列表。
</details>

<details>
<summary><b>字数要求不是 7000 字怎么办？</b></summary>

直接跟 Claude 说："报告字数要求 5000 字左右"，它会自动调整。
</details>

---

## 版本历史

| 版本 | 日期 | 变更 |
|------|------|------|
| **v2.0** | 2026-03-14 | 大版本升级：新增 6 个函数、合并文档、完善模板结构 |
| v1.0 | 2026-01-01 | 初始发布 |

## License

MIT

## 致谢

- [python-docx](https://python-docx.readthedocs.io/) - Word 文档操作库
- [Claude Code](https://docs.anthropic.com/en/docs/claude-code) - Anthropic 官方 CLI

---

<div align="center">

**如果这个工具帮到了你，点个 Star 吧** ⭐

*一个被课设报告折磨过的人，写给正在被折磨的你。*

</div>
