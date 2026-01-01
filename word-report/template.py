# -*- coding: utf-8 -*-
"""
项目报告生成脚本
字数要求：6800字左右
"""
from docx import Document
from docx.shared import Pt, Cm, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement
from docx.oxml import parse_xml
import os
import copy

OUTPUT_PATH = r"E:\report\word\项目报告.docx"
TEMPLATE_PATH = r"C:\Users\31930\.claude\skills\word-report\examples\课程设计报告书模板.docx"

# 封面信息
COVER_INFO = {
    "school": "学校名称",
    "college": "学院名称",
    "course": "课程名称",
    "title": "项目标题",
    "major_class": "专业班级",
    "name": "姓名",
    "student_id": "学号",
    "teacher": "授课教师",
    "year": "年份",
    "month": "",
    "day": ""
}
# ================================================


def set_chinese_font(run, font_name='宋体', font_size=Pt(10.5), bold=False, italic=False, underline=False):
    """设置中文字体"""
    run.font.name = font_name
    run.font.size = font_size
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = underline
    # 设置中文字体
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), font_name)


def set_paragraph_format(paragraph, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                         first_line_indent=None, line_spacing=1.25,
                         space_before=Pt(0), space_after=Pt(0)):
    """设置段落格式"""
    pf = paragraph.paragraph_format
    pf.alignment = alignment
    pf.line_spacing = line_spacing
    pf.space_before = space_before
    pf.space_after = space_after
    if first_line_indent is not None:
        pf.first_line_indent = first_line_indent


def enable_trailing_space_underline(doc):
    """开启'为尾部空格添加下划线'设置"""
    settings = doc.settings.element
    compat = settings.find(qn('w:compat'))
    if compat is None:
        compat = OxmlElement('w:compat')
        settings.append(compat)
    ul_trail_space = OxmlElement('w:ulTrailSpace')
    compat.append(ul_trail_space)


def create_cover_page(doc, info):
    """创建封面页"""
    # 学校名称
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.5)
    run = p.add_run(info["school"])
    set_chinese_font(run, font_size=Pt(22), bold=True)

    # 学院名称
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.5)
    run = p.add_run(info["college"])
    set_chinese_font(run, font_size=Pt(18), bold=True)

    # 空行
    doc.add_paragraph()
    doc.add_paragraph()

    # 课程名称
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.5)
    run = p.add_run("课程名称：")
    set_chinese_font(run, font_size=Pt(14))
    run = p.add_run("     " + info["course"])
    set_chinese_font(run, font_size=Pt(14), underline=True)

    # 空行
    doc.add_paragraph()

    # 封面表格
    table = doc.add_table(rows=6, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 设置表格宽度
    table.columns[0].width = Cm(3)
    table.columns[1].width = Cm(6)

    # 表格内容
    table_data = [
        ("题    目：", info["title"]),
        ("专业班级：", info["major_class"]),
        ("姓    名：", info["name"]),
        ("学    号：", info["student_id"]),
        ("授课教师：", info["teacher"]),
        ("成    绩：", "")
    ]

    for i, (label, value) in enumerate(table_data):
        # 标签单元格
        cell_label = table.cell(i, 0)
        p = cell_label.paragraphs[0]
        set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.RIGHT, line_spacing=1.5)
        run = p.add_run(label)
        set_chinese_font(run, font_size=Pt(14))

        # 值单元格（带下划线）
        cell_value = table.cell(i, 1)
        p = cell_value.paragraphs[0]
        set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.5)
        # 使用空格填充以显示下划线
        value_text = value if value else "          "
        value_text = value_text.ljust(20)  # 确保有足够的空格显示下划线
        run = p.add_run(value_text)
        set_chinese_font(run, font_size=Pt(14), underline=True)

    # 空行
    doc.add_paragraph()
    doc.add_paragraph()

    # 日期
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.5)
    date_text = f"{info['year']} 年    月    日"
    run = p.add_run(date_text)
    set_chinese_font(run, font_size=Pt(14))

    # 分页符
    doc.add_page_break()


def add_chapter_title(doc, text):
    """添加一级标题（章标题）"""
    p = doc.add_paragraph(text, style='Heading 1')
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                         space_before=Pt(0), space_after=Pt(4), line_spacing=1.25)
    for run in p.runs:
        set_chinese_font(run, font_size=Pt(15), bold=True)
    return p


def add_section_title(doc, text):
    """添加二级标题（小节标题）"""
    p = doc.add_paragraph(text, style='Heading 2')
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                         space_before=Pt(2), space_after=Pt(0), line_spacing=1.25)
    for run in p.runs:
        set_chinese_font(run, font_size=Pt(12), bold=True)
    return p


def add_summary_paragraph(doc, text):
    """添加总领段落（无缩进）"""
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                         first_line_indent=Cm(0), line_spacing=1.25,
                         space_before=Pt(0), space_after=Pt(0))
    run = p.add_run(text)
    set_chinese_font(run, font_size=Pt(10.5))
    return p


def add_body_paragraph(doc, text):
    """添加正文段落（首行缩进2字符）"""
    p = doc.add_paragraph()
    # 2字符缩进，五号字体约为0.74cm
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                         first_line_indent=Cm(0.74), line_spacing=1.25,
                         space_before=Pt(0), space_after=Pt(0))
    run = p.add_run(text)
    set_chinese_font(run, font_size=Pt(10.5))
    return p


def add_code_block(doc, code_text):
    """添加代码块"""
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                         first_line_indent=Cm(0), line_spacing=1.0,
                         space_before=Pt(3), space_after=Pt(3))
    run = p.add_run(code_text)
    set_chinese_font(run, font_name='宋体', font_size=Pt(9))
    return p


def add_image_placeholder(doc, description, fig_num):
    """添加图片占位符"""
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.25)
    run = p.add_run(f"[{description} - 图{fig_num}]")
    set_chinese_font(run, font_size=Pt(10.5), italic=True)
    return p


def add_references_title(doc):
    """添加参考文献标题"""
    # 分页符
    doc.add_page_break()

    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                         space_before=Pt(0), space_after=Pt(4), line_spacing=1.25)
    run = p.add_run("参考文献")
    set_chinese_font(run, font_size=Pt(15), bold=True)
    return p


def add_reference_item(doc, text):
    """添加参考文献条目"""
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                         first_line_indent=Cm(0), line_spacing=1.25,
                         space_before=Pt(0), space_after=Pt(0))
    run = p.add_run(text)
    set_chinese_font(run, font_size=Pt(10.5))
    return p


def remove_content_after_toc(doc, toc_end_index=35):
    """删除目录之后的所有段落，保留所有表格"""
    # 只删除段落，不删除表格（保留封面表格和评分表）
    for i in range(len(doc.paragraphs) - 1, toc_end_index, -1):
        para = doc.paragraphs[i]
        para._element.getparent().remove(para._element)


def move_table_to_end(doc, table_index):
    """将指定表格移到文档末尾"""
    table = doc.tables[table_index]
    table_element = table._element

    # 从原位置移除
    table_element.getparent().remove(table_element)

    # 添加到文档末尾
    doc._element.body.append(table_element)


def create_sample_report():
    """创建项目报告 - 字数约6800字"""
    print("正在读取模板文档...")
    doc = Document(TEMPLATE_PATH)

    # 步骤1：删除目录之后的内容
    print("正在清理模板正文...")
    remove_content_after_toc(doc, toc_end_index=35)

    # 添加分页符，开始正文
    doc.add_page_break()

    # 步骤2：用代码创建正文内容
    print("正在生成正文内容...")

    # ========== 第1章 设计题目 ==========
    add_chapter_title(doc, "第1章 设计题目")
    add_body_paragraph(doc, "项目简介背景段落，描述项目来源和应用场景。")
    add_body_paragraph(doc, "系统架构概述段落，描述技术栈和整体架构。")
    add_body_paragraph(doc, "核心特色功能段落，描述项目的主要特点和创新点。")
    doc.add_paragraph()

    # ========== 第2章 开发环境 ==========
    add_chapter_title(doc, "第2章 开发环境")
    add_section_title(doc, "2.1 硬件环境")
    add_body_paragraph(doc, "硬件环境描述，包括处理器、内存、硬盘、测试设备等配置信息。")
    add_section_title(doc, "2.2 软件环境")
    add_body_paragraph(doc, "软件环境描述，包括操作系统、开发工具、服务器、数据库等软件配置。")
    doc.add_paragraph()

    # ========== 第3章 开发工具 ==========
    add_chapter_title(doc, "第3章 开发工具")
    add_section_title(doc, "3.1 开发工具1")
    add_body_paragraph(doc, "开发工具1介绍，描述其功能、特点和在项目中的应用。")
    add_section_title(doc, "3.2 开发工具2")
    add_body_paragraph(doc, "开发工具2介绍，描述其功能、特点和在项目中的应用。")
    add_section_title(doc, "3.3 开发工具3")
    add_body_paragraph(doc, "开发工具3介绍，描述其功能、特点和在项目中的应用。")
    add_section_title(doc, "3.4 开发工具4")
    add_body_paragraph(doc, "开发工具4介绍，描述其功能、特点和在项目中的应用。")
    add_section_title(doc, "3.5 开发工具5")
    add_body_paragraph(doc, "开发工具5介绍，描述其功能、特点和在项目中的应用。")
    doc.add_paragraph()

    # ========== 第4章 完成时间 ==========
    add_chapter_title(doc, "第4章 完成时间")
    add_body_paragraph(doc, "项目开发周期描述，包括起止时间、各阶段工作内容等。")
    doc.add_paragraph()

    # ========== 第5章 设计思想 ==========
    add_chapter_title(doc, "第5章 设计思想")
    add_section_title(doc, "5.1 架构设计")
    add_body_paragraph(doc, "架构设计段落1，描述整体架构和各层职责。")
    add_body_paragraph(doc, "架构设计段落2，描述具体设计模式和实现方式。")
    add_section_title(doc, "5.2 设计模式或架构风格")
    add_body_paragraph(doc, "设计模式或架构风格描述。")
    add_section_title(doc, "5.3 关键技术或策略")
    add_body_paragraph(doc, "关键技术或策略描述。")
    add_section_title(doc, "5.4 核心机制")
    add_body_paragraph(doc, "核心机制描述。")
    doc.add_paragraph()

    # ========== 第6章 设计过程及设计步骤 ==========
    add_chapter_title(doc, "第6章 设计过程及设计步骤")
    add_section_title(doc, "6.1 数据库设计")
    add_body_paragraph(doc, "数据库设计阶段介绍，包括概念设计、逻辑设计、物理设计。")
    add_body_paragraph(doc, "具体表结构描述，列出主要数据表和字段说明。")
    add_image_placeholder(doc, "数据库ER图", 1)
    add_section_title(doc, "6.2 功能模块开发")
    add_body_paragraph(doc, "功能模块开发概述。")
    add_body_paragraph(doc, "前端开发详细描述。")
    add_body_paragraph(doc, "后端开发详细描述。")
    add_body_paragraph(doc, "特色功能实现描述。")
    add_image_placeholder(doc, "系统功能模块图", 2)
    add_section_title(doc, "6.3 接口联调")
    add_body_paragraph(doc, "接口联调过程描述，包括测试方法、遇到的问题和解决方案。")
    doc.add_paragraph()

    # ========== 第7章 测试运行 ==========
    add_chapter_title(doc, "第7章 测试运行")
    add_section_title(doc, "7.1 功能测试")
    add_body_paragraph(doc, "功能测试范围和方法描述。")
    add_body_paragraph(doc, "具体测试过程和结果描述。")
    add_image_placeholder(doc, "运行截图1", 3)
    add_image_placeholder(doc, "运行截图2", 4)
    add_section_title(doc, "7.2 性能测试")
    add_body_paragraph(doc, "性能测试指标和结果描述。")
    add_section_title(doc, "7.3 运行结果")
    add_body_paragraph(doc, "系统运行情况和效果描述。")
    add_image_placeholder(doc, "运行截图3", 5)
    doc.add_paragraph()

    # ========== 第8章 评价与修订 ==========
    add_chapter_title(doc, "第8章 评价与修订")
    add_section_title(doc, "8.1 功能评价")
    add_body_paragraph(doc, "功能完成情况评价。")
    add_section_title(doc, "8.2 技术评价")
    add_body_paragraph(doc, "技术实现评价。")
    add_section_title(doc, "8.3 存在问题")
    add_body_paragraph(doc, "系统存在的不足和问题。")
    add_section_title(doc, "8.4 改进建议")
    add_body_paragraph(doc, "后续改进方向和建议。")
    doc.add_paragraph()

    # ========== 第9章 设计体会 ==========
    add_chapter_title(doc, "第9章 设计体会")
    add_body_paragraph(doc, "技术和能力提升收获。")
    add_body_paragraph(doc, "核心技术学习体会。")
    add_body_paragraph(doc, "问题解决经历和心得。")
    add_body_paragraph(doc, "工程思维认识提升。")
    add_body_paragraph(doc, "团队合作体会（如适用）。")
    add_body_paragraph(doc, "不足之处和未来展望。")
    doc.add_paragraph()

    # ========== 参考文献 ==========
    add_references_title(doc)
    add_reference_item(doc, "[1] 作者. 书名[M]. 出版社: 年份.")
    add_reference_item(doc, "[2] 作者. 书名[M]. 出版社: 年份.")
    add_reference_item(doc, "[3] 作者. 书名[M]. 出版社: 年份.")
    add_reference_item(doc, "[4] 作者. 书名[M]. 出版社: 年份.")
    add_reference_item(doc, "[5] 作者. 书名[M]. 出版社: 年份.")
    add_reference_item(doc, "[6] 作者. 文档名[EB/OL]. URL, 年份.")
    add_reference_item(doc, "[7] 作者. 文档名[EB/OL]. URL, 年份.")
    add_reference_item(doc, "[8] 作者. 文档名[EB/OL]. URL, 年份.")

    # 添加分页符，确保评分表在新页开始
    doc.add_page_break()

    # 将评分表移到文档末尾
    print("正在移动评分表到文档末尾...")
    if len(doc.tables) >= 2:
        move_table_to_end(doc, 1)  # 移动第二个表格（评分表）到末尾

    # 保存文档
    print("正在保存文档...")
    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    doc.save(OUTPUT_PATH)
    print(f"✓ 报告已生成：{OUTPUT_PATH}")


def main():
    """主程序入口"""
    print("正在生成项目报告...")
    create_sample_report()
    print("完成！")


if __name__ == "__main__":
    main()
