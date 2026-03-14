# -*- coding: utf-8 -*-
"""
模板标题提取工具

用途：当用户提供了参考文档时，提取其中的章节标题（一级标题、二级标题），
      用于确定报告的章节结构。

使用方法：
    python analyze_template.py "D:\path\to\template.docx"
"""
from docx import Document
import sys


def analyze_template(template_path):
    """提取模板中的章节标题

    Args:
        template_path: Word 模板文件路径

    输出信息：
        - 一级标题（章标题）
        - 二级标题（小节标题）
    """
    doc = Document(template_path)

    print("=" * 60)
    print("📋 参考文档标题提取")
    print("=" * 60)

    chapter_titles = []
    section_titles = []

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        style_name = para.style.name if para.style else ""

        # 方式1：通过 Heading 样式识别标题
        if style_name.startswith('Heading 1') or style_name == '标题 1':
            if text and text not in chapter_titles:
                chapter_titles.append(text)
                print(f"  [章-样式] {text}")
                continue
        elif style_name.startswith('Heading 2') or style_name == '标题 2':
            if text and text not in section_titles:
                section_titles.append(text)
                print(f"  [节-样式] {text}")
                continue

        # 方式2：通过文本前缀识别（兜底）
        # 识别一级标题（以"第X章"开头）
        if text.startswith("第") and "章" in text[:6]:
            if text not in chapter_titles:
                chapter_titles.append(text)
                print(f"  [章] {text}")
        elif len(text) >= 3 and text[0].isdigit() and text[1] == " " and not text[2].isdigit():
            if text not in chapter_titles:
                chapter_titles.append(text)
                print(f"  [章] {text}")

        # 识别二级标题（数字+点+数字+空格，如"1.1"）
        elif len(text) >= 4 and text[0].isdigit() and text[1] == "." and text[2].isdigit():
            if text not in section_titles:
                section_titles.append(text)
                print(f"  [节] {text}")

    print("\n" + "=" * 60)
    print(f"✅ 共提取 {len(chapter_titles)} 个一级标题")
    print(f"✅ 共提取 {len(section_titles)} 个二级标题")
    print("=" * 60)

    return chapter_titles, section_titles


def main():
    """命令行入口"""
    if len(sys.argv) < 2:
        print("用法: python analyze_template.py <模板路径>")
        print("示例: python analyze_template.py \"D:\\templates\\报告模板.docx\"")
        sys.exit(1)

    template_path = sys.argv[1]
    analyze_template(template_path)


if __name__ == "__main__":
    main()
