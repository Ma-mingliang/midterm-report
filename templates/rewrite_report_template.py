"""
中期考核报告生成脚本模板 v4.0（通用版）
适用于任意研究方向，通过动态模块结构生成报告

使用方法：
1. 修改 BASE_DIR 为实际工作目录
2. 修改 modules 列表，定义各研究模块
3. 修改 write_section1/2/3 函数中的实际内容
4. 运行: E:/Anaconda/python.exe rewrite_report.py

注意：
- 不要使用 conda run -n base python -c "多行代码"（Windows不支持）
- 不要信任图片文件名，必须用视觉模型核验图片实际内容
- 公式使用 pandoc LaTeX → OMML 转换，确保渲染正确（需要安装 pandoc）

环境要求：
- Python >= 3.7
- 依赖：python-docx, PyMuPDF (fitz), Pillow, lxml
"""

from __future__ import annotations

import os
import shutil
import sys
from dataclasses import dataclass, field
from typing import Optional

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from lxml import etree
from lxml.etree import _Element

# ============================================================
# 配置区域 - 根据实际情况修改
# ============================================================
BASE_DIR: str = r'工作目录路径'
WRITE_FILE: str = os.path.join(BASE_DIR, '写入文件.docx')
IMG_DIR: str = os.path.join(BASE_DIR, '图片')  # 用户提供的图片（可选）


# ============================================================
# 模块定义 - 动态结构
# ============================================================
@dataclass
class Module:
    """研究模块定义"""
    title: str           # 模块标题
    summary: str         # 1-2句话的概述（用于第二层）
    description: str     # 详细描述（用于第三层）
    images: list[str] = field(default_factory=list)  # 图片文件名列表
    formulas: list[str] = field(default_factory=list)  # 公式标识列表


# ============================================================
# 动态模块列表 - 根据实际研究内容修改
# ============================================================
# 示例：以下是通用模板，请根据PPT/论文实际内容替换
modules: list[Module] = [
    Module(
        title='数据采集与预处理',
        summary='（1）[模块名称]，[做了什么]，[怎么做的]，[达到什么效果/为后续提供什么]。',
        description='详细描述该模块的方法、实验和结果...',
        images=['图2.png', '图3.png'],
        formulas=[],
    ),
    Module(
        title='核心方法设计',
        summary='（2）[模块名称]，[做了什么]，[怎么做的]，[达到什么效果/为后续提供什么]。',
        description='详细描述该模块的方法、实验和结果...',
        images=['图4.png', '图5.png', '图6.png'],
        formulas=['loss'],
    ),
    Module(
        title='实验验证与分析',
        summary='（3）[模块名称]，[做了什么]，[怎么做的]，[达到什么效果/为后续提供什么]。',
        description='详细描述该模块的方法、实验和结果...',
        images=['图7.png', '图8.png'],
        formulas=[],
    ),
    # 根据实际需要添加更多模块...
]


# ============================================================
# 工具函数
# ============================================================
def _apply_spacing(pPr: _Element, line: str = '360', after: str = '0') -> None:
    """统一设置段落行距和段后间距"""
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = etree.SubElement(pPr, qn('w:spacing'))
    spacing.set(qn('w:line'), line)
    spacing.set(qn('w:lineRule'), 'auto')
    spacing.set(qn('w:after'), after)


def _apply_indent(pPr: _Element, first_indent: Optional[str] = None, left_indent: Optional[str] = None) -> None:
    """统一设置段落缩进"""
    if first_indent is None and left_indent is None:
        return
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = etree.SubElement(pPr, qn('w:ind'))
    if first_indent is not None:
        ind.set(qn('w:firstLine'), first_indent)
    if left_indent is not None:
        ind.set(qn('w:left'), left_indent)


def set_run_font(
    run,
    font_name: str = 'Times New Roman',
    font_size: Pt = Pt(12),
    bold: Optional[bool] = None,
    east_asia: Optional[str] = '宋体',
) -> None:
    """设置混合字体：中文宋体 + 英文 Times New Roman"""
    run.font.name = font_name
    run.font.size = font_size
    if bold is not None:
        run.bold = bold
    r = run._element
    rPr = r.find(qn('w:rPr'))
    if rPr is None:
        rPr = r.makeelement(qn('w:rPr'), {})
        r.insert(0, rPr)
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rPr.makeelement(qn('w:rFonts'), {})
        rPr.insert(0, rFonts)
    if east_asia is not None:
        rFonts.set(qn('w:eastAsia'), east_asia)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)


def clear_cell(cell) -> None:
    """清空表格单元格内容"""
    for i in range(len(cell.paragraphs) - 1, 0, -1):
        p = cell.paragraphs[i]
        p._element.getparent().remove(p._element)
    first_p = cell.paragraphs[0]
    for run in first_p.runs:
        run.text = ''
    for r in first_p._element.findall(qn('w:r')):
        first_p._element.remove(r)


def add_para(
    cell,
    text: str,
    font_name: str = 'Times New Roman',
    font_size: Pt = Pt(12),
    bold: Optional[bool] = None,
    alignment=None,
    first_indent: Optional[str] = '480',
    left_indent: Optional[str] = '34',
    east_asia: Optional[str] = '宋体',
):
    """向单元格添加段落，自动设置1.5倍行距和首行缩进"""
    first_p = cell.paragraphs[0]
    if len(cell.paragraphs) == 1 and not first_p.text.strip() and len(first_p._element.findall(qn('w:r'))) == 0:
        p = first_p
    else:
        p = cell.add_paragraph()
    if alignment is not None:
        p.alignment = alignment
    run = p.add_run(text)
    set_run_font(run, font_name, font_size, bold, east_asia=east_asia)
    pPr = p._element.get_or_add_pPr()
    _apply_spacing(pPr)
    _apply_indent(pPr, first_indent, left_indent)
    return p


def add_img(cell, img_path: str, width=Inches(5.0)):
    """向单元格添加图片（居中）

    Raises:
        FileNotFoundError: 图片文件不存在时抛出
    """
    if not os.path.exists(img_path):
        raise FileNotFoundError(
            f'图片文件不存在: {img_path}\n'
            f'请检查图片路径是否正确，或确认图片已从PPT/论文中提取。'
        )
    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(img_path, width=width)
    return p


def add_caption(cell, text: str, font_size: Pt = Pt(10)):
    """添加图片图注（居中，带a3样式，首行缩进400，1.5倍行距）"""
    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    set_run_font(run, 'Times New Roman', font_size, east_asia=None)
    pPr = p._element.get_or_add_pPr()
    # 段落样式：a3（与模板图注一致）
    pStyle = pPr.find(qn('w:pStyle'))
    if pStyle is None:
        pStyle = etree.SubElement(pPr, qn('w:pStyle'))
    pStyle.set(qn('w:val'), 'a3')
    _apply_spacing(pPr)
    _apply_indent(pPr, first_indent='400')
    return p


def img(name: str) -> Optional[str]:
    """图片路径查找（优先级：图片/ → ppt_images/ → paper_images/）

    Returns:
        图片文件的完整路径

    Raises:
        FileNotFoundError: 所有路径都找不到图片时抛出，附带详细诊断信息
    """
    search_paths: list[tuple[str, str]] = [
        (IMG_DIR, '用户提供的图片目录'),
        (os.path.join(BASE_DIR, 'ppt_images'), 'PPT提取的图片'),
        (os.path.join(BASE_DIR, 'paper_images'), '论文提取的图片'),
    ]

    for directory, desc in search_paths:
        candidate = os.path.join(directory, name)
        if os.path.exists(candidate):
            return candidate

    # 构建诊断信息
    checked = '\n'.join(f'  - {desc}: {os.path.join(d, name)}' for d, desc in search_paths)
    raise FileNotFoundError(
        f'图片 "{name}" 在以下路径中均未找到:\n{checked}\n\n'
        f'请确认:\n'
        f'  1. 图片已从PPT/论文中提取（运行阶段2的提取脚本）\n'
        f'  2. 图片文件名正确（不要信任文件名，用视觉模型核验）\n'
        f'  3. 如使用用户图片目录，确认 图片/ 目录存在且包含该文件'
    )


# ============================================================
# 章节写入函数
# ============================================================
def write_section1(cell, transcript_data: Optional[dict] = None) -> None:
    """一、思想品德与业务学习情况自述

    Args:
        cell: 表格单元格
        transcript_data: 成绩单数据（可选），包含：
            - total_credits: 总学分
            - gpa: 绩点
            - courses: 课程列表 [{'name': '课程名', 'score': 分数, 'category': '类别'}]
            - has_remake: 是否有重修
    """
    clear_cell(cell)

    # 思想政治
    add_para(cell, '本人自[入学年月]以来，严格遵守学校的各项规章制度，'
             '在思想上积极向党组织靠拢，认真学习马克思主义基本理论，'
             '拥护党的路线方针政策，坚持四项基本原则。')

    # 学习态度+成绩总述
    if transcript_data:
        credits = transcript_data.get('total_credits', 'XX')
        gpa = transcript_data.get('gpa', 'X.XX')
        remake = '无任何重修课程' if not transcript_data.get('has_remake', False) else ''
        add_para(cell, f'在学习上，本人勤奋努力，积极向上，学习成绩优良。'
                 f'目前培养计划总学分{credits}分，现已圆满修满全部学分，'
                 f'总平均绩点达{gpa}，{remake}。')
    else:
        add_para(cell, '在学习上，本人勤奋努力，积极向上，学习成绩优良。'
                 '（请根据成绩单填写具体学分和绩点）')

    # 核心课程（lyw-模板：逐门列出）
    add_para(cell, '在核心课程方面，（请根据成绩单逐门列出课程名称、分数和收获）。')

    # 科研能力
    add_para(cell, '在科研上，本人认真钻研，积极探索，在导师的悉心指导下，'
             '本人参与了多项科研项目，掌握了[相关技能]，具备了独立开展科研工作的能力。')

    # 学术道德
    add_para(cell, '在学术道德方面，本人严格遵守学术规范和道德准则，'
             '在相关课程的指导下，树立了正确的学术价值观。')

    # 身体素质
    add_para(cell, '在身体素质方面，本人平时积极参加体育锻炼，保持良好的身心状态，'
             '为科研工作提供了坚实的基础。')


def write_section2(cell) -> None:
    """二、已完成的科研工作

    动态生成：根据 modules 列表自动展开各模块。
    三层行文逻辑：
    1. 开篇总述：研究目的 + 整体框架 + 框架图
    2. N部分简要概述：每个模块1-2句话概括
    3. 详细展开："已完成工作：" + 每个模块详细描述
    """
    clear_cell(cell)

    # ============================================================
    # 第一层：开篇总述
    # ============================================================
    add_para(cell, '本文的研究目的是[从PPT首页/摘要提取研究目的]。'
             f'如图1所示为本文的整体研究框架。')
    add_img(cell, img('图1.png'), width=Inches(5.5))
    add_caption(cell, '图1 整体研究框架')

    # ============================================================
    # 第二层：N部分简要概述
    # ============================================================
    add_para(cell, f'目前已完成的工作主要包括以下{len(modules)}个方面：')

    for i, module in enumerate(modules, 1):
        # 使用模块的 summary，如果没有则生成默认格式
        summary = module.summary if module.summary else f'（{i}）{module.title}。'
        add_para(cell, summary)

    # ============================================================
    # 第三层：详细展开
    # ============================================================
    add_para(cell, '已完成工作：')

    for i, module in enumerate(modules, 1):
        # 模块标题
        add_para(cell, f'（{i}）{module.title}。')

        # 模块详细描述
        add_para(cell, module.description)

        # 插入模块图片
        for img_name in module.images:
            add_img(cell, img(img_name), width=Inches(5.0))
            # 图注从文件名推断，实际应根据交叉验证结果填写
            caption = img_name.replace('.png', '').replace('.jpg', '')
            add_caption(cell, f'{caption}')

        # 插入模块公式
        # from omml_formulas import create_formula_paragraph
        # for formula_id in module.formulas:
        #     create_formula_paragraph(cell, formula_id)

    # 总结
    add_para(cell, '综上所述，本人已完成[总结已完成的工作内容]。')


def write_section3(cell) -> None:
    """三、下一步科研计划

    要求：
    - 分3个时间段，每段包含明确的日期范围
    - 每段描述2-3个具体研究任务
    - 与第二章已完成工作逻辑衔接
    - 包含论文撰写里程碑
    """
    clear_cell(cell)

    add_para(cell, '根据目前的研究进展和课题计划，下一步的科研工作主要包括以下三个阶段：')
    add_para(cell, '20XX.XX-20XX.XX：[第一阶段内容，与已完成工作衔接]...', first_indent='440')
    add_para(cell, '20XX.XX-20XX.XX：[第二阶段内容，深化研究]...', first_indent='440')
    add_para(cell, '20XX.XX-20XX.XX：[第三阶段内容，论文撰写与答辩准备]...', first_indent='440')


# ============================================================
# 主函数
# ============================================================
def main() -> None:
    # 检查工作目录
    if not os.path.isdir(BASE_DIR):
        print(f'错误：工作目录不存在: {BASE_DIR}')
        print('请修改脚本中的 BASE_DIR 为实际工作目录路径。')
        sys.exit(1)

    # 从模板恢复（模板文件已内置，使用工作目录中的模板或创建新文档）
    template_path = os.path.join(BASE_DIR, '模板.docx')
    if os.path.exists(template_path):
        shutil.copy2(template_path, WRITE_FILE)
        print(f'从模板恢复: {template_path}')
    else:
        doc = Document()
        doc.save(WRITE_FILE)
        print(f'未找到模板文件，已创建新文档: {WRITE_FILE}')

    doc = Document(WRITE_FILE)

    if len(doc.tables) < 1:
        print(f'错误：文档中没有表格。请确认模板文件包含正确的表格结构。')
        print(f'当前文档: {WRITE_FILE}')
        sys.exit(1)

    table = doc.tables[0]
    if len(table.rows) < 6:
        print(f'错误：表格行数不足（需要至少6行，当前{len(table.rows)}行）。')
        print(f'请确认模板文件的表格结构正确。')
        sys.exit(1)

    print('Writing Section 1...')
    write_section1(table.rows[1].cells[0])

    print('Writing Section 2...')
    write_section2(table.rows[3].cells[0])

    print('Writing Section 3...')
    write_section3(table.rows[5].cells[0])

    doc.save(WRITE_FILE)
    print(f'Document saved: {WRITE_FILE}')


if __name__ == '__main__':
    main()
