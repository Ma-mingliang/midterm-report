"""
中期考核报告生成脚本（通用模板）

使用方法：
1. 将本文件复制到你的工作目录
2. 修改下方【用户配置区】的路径和文件名
3. 根据你的研究内容填写 MODULES、FORMULAS、fig_captions
4. 运行: python rewrite_report.py

依赖：python-docx, lxml, pandoc, pdflatex
"""

import os
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, field

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from lxml import etree

# ============================================================
# 【用户配置区】根据实际情况修改
# ============================================================

# 工作目录（脚本所在目录）
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 模板文件名（你的格式模板，含封面页和表格结构）
TEMPLATE_FILE = os.path.join(BASE_DIR, '模板.docx')

# 输出文件名
WRITE_FILE = os.path.join(BASE_DIR, '写入文件.docx')

# 图片目录（所有待插入报告的图片）
IMG_DIR = os.path.join(BASE_DIR, '图片')


# ============================================================
# 公式定义（LaTeX 格式）
# ============================================================
# 按需添加你的公式，key 为公式名称（在 MODULES 中引用），value 为 LaTeX 字符串
FORMULAS: dict[str, str] = {
    # 示例：
    # 'loss': r'L_{total} = \alpha L_1 + (1 - \alpha) L_2',
    # 'eq1':  r'y = f(x) = \sum_{i=1}^{n} w_i x_i',
}


# ============================================================
# 图片路径
# ============================================================
def img(name: str) -> str:
    """根据文件名返回图片完整路径（从 图片/ 目录查找）"""
    path = os.path.join(IMG_DIR, name)
    if not os.path.exists(path):
        raise FileNotFoundError(f'图片不存在: {path}')
    return path


# ============================================================
# 工具函数（一般不需要修改）
# ============================================================
def set_run_font(run, font_name='Times New Roman', font_size=Pt(12), bold=None, east_asia='宋体'):
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


def clear_cell(cell):
    """清空表格单元格内容"""
    for i in range(len(cell.paragraphs) - 1, 0, -1):
        p = cell.paragraphs[i]
        p._element.getparent().remove(p._element)
    first_p = cell.paragraphs[0]
    for run in first_p.runs:
        run.text = ''
    for r in first_p._element.findall(qn('w:r')):
        first_p._element.remove(r)


def _apply_spacing(pPr):
    """设置行距：1.5倍，段后间距0"""
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = etree.SubElement(pPr, qn('w:spacing'))
    spacing.set(qn('w:line'), '360')
    spacing.set(qn('w:lineRule'), 'auto')
    spacing.set(qn('w:after'), '0')
    return spacing


def _apply_indent(pPr, first_indent='480', left_indent='34', first_line_chars='200'):
    """设置首行缩进和左缩进"""
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = etree.SubElement(pPr, qn('w:ind'))
    if first_line_chars is not None:
        ind.set(qn('w:firstLineChars'), first_line_chars)
    if first_indent is not None:
        ind.set(qn('w:firstLine'), first_indent)
    if left_indent is not None:
        ind.set(qn('w:left'), left_indent)


def add_para(cell, text, font_name='Times New Roman', font_size=Pt(12), bold=None,
             alignment=None, first_indent='480', left_indent='34', east_asia='宋体'):
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
    if alignment is None:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if first_indent is not None or left_indent is not None:
        _apply_indent(pPr, first_indent, left_indent)
    return p


def add_img(cell, img_path, width=Inches(5.0)):
    """向单元格添加图片（居中）"""
    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(img_path, width=width)
    pPr = p._element.get_or_add_pPr()
    _apply_spacing(pPr)
    return p


def add_caption(cell, text, font_size=Pt(10)):
    """添加图片图注（居中，黑体，首行缩进400，1.5倍行距）"""
    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    set_run_font(run, 'Times New Roman', font_size, east_asia='黑体')
    pPr = p._element.get_or_add_pPr()
    pStyle = pPr.find(qn('w:pStyle'))
    if pStyle is None:
        pStyle = etree.SubElement(pPr, qn('w:pStyle'))
    pStyle.set(qn('w:val'), 'a3')
    _apply_spacing(pPr)
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = etree.SubElement(pPr, qn('w:ind'))
    ind.set(qn('w:firstLine'), '400')
    return p


def add_formula(cell, formula_name: str):
    """
    通过 pdflatex + pandoc 将 LaTeX 公式渲染为 Word 原生 OMML 格式。
    流程：LaTeX → pdflatex 编译 → pandoc 转换 docx → 提取 oMath 元素 → 插入单元格
    """
    latex = FORMULAS.get(formula_name)
    if latex is None:
        print(f'警告: 公式 "{formula_name}" 未在 FORMULAS 中定义')
        return
    tmp_dir = BASE_DIR
    with tempfile.NamedTemporaryFile(mode='w', suffix='.tex', delete=False, encoding='utf-8', dir=tmp_dir) as f:
        f.write(f'\\documentclass{{article}}\\begin{{document}}${latex}$\\end{{document}}')
        tex_path = f.name
    docx_path = tex_path.replace('.tex', '.docx')
    log_path = tex_path.replace('.tex', '.log')
    aux_path = tex_path.replace('.tex', '.aux')
    pdf_path = tex_path.replace('.tex', '.pdf')
    try:
        subprocess.run(
            ['pdflatex', '-interaction=nonstopmode', tex_path],
            capture_output=True, timeout=30, cwd=tmp_dir
        )
        subprocess.run(
            ['pandoc', tex_path, '-o', docx_path, '--from=latex', '--to=docx'],
            check=True, capture_output=True, timeout=30
        )
        temp_doc = Document(docx_path)
        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pPr = p._element.get_or_add_pPr()
        _apply_spacing(pPr)
        found = False
        for temp_p in temp_doc.paragraphs:
            for elem in temp_p._element:
                if elem.tag.endswith('}oMath') or elem.tag.endswith('}oMathPara'):
                    p._element.append(elem)
                    found = True
        if not found:
            print(f'公式 {formula_name}: 未找到oMath元素，尝试直接渲染')
    except Exception as e:
        print(f'公式渲染失败 {formula_name}: {e}')
        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f'[{formula_name}]')
        set_run_font(run)
    finally:
        for fp in [tex_path, docx_path, log_path, aux_path, pdf_path]:
            if os.path.exists(fp):
                os.unlink(fp)
    return p


# ============================================================
# Module 数据结构
# ============================================================
@dataclass
class Module:
    """
    研究模块定义。

    字段说明：
    - title: 模块标题（如"多模态数据采集与预处理"）
    - summary: 概述层的1-2句话概括（如"（1）多模态数据采集与预处理，搭建采集平台..."）
    - description: 详细展开的正文，用 \\n 分隔段落
    - images: 图片文件名列表（从 图片/ 目录读取）
    - image_after: {图片索引: 段落索引}，表示第N张图片放在第M段之后
    - formula_after: {段落索引: 公式名称}，表示第M段之后插入指定公式
    """
    title: str
    summary: str
    description: str
    images: list[str]
    image_after: dict[int, int] = field(default_factory=dict)
    formula_after: dict[int, str] = field(default_factory=dict)


# ============================================================
# 【用户配置区】模块定义
# ============================================================
# 根据你的研究内容定义模块。每个 Module 对应第二章的一个技术模块。
# 模块数量一般为 4-8 个，根据 PPT 中的实际模块数确定。
#
# description 写作规则：
# 1. 用 \n 分隔段落（每个 \n = 一个新段落）
# 2. 正文中不要写数学公式（如 y=f(x)），公式通过 add_formula 渲染
# 3. 正文中引用图片用"如图X所示"，图片通过 image_after 在对应段落后插入
# 4. 公式引入语必须明确（如"定义为："、"表示为："），不能凭空出现公式
#
MODULES = [
    # ---- 示例模块（替换为你的实际内容）----
    Module(
        title='模块标题',
        summary='（1）模块标题，做了什么，怎么做的，达到什么效果。',
        description=(
            '第一段：背景与目的。为了实现XXX，本人设计了...\n'
            '第二段：方法描述。该方法的核心思想是...定义为：\n'
            '第三段：实验验证。如图X所示为实验结果...'
        ),
        images=['图01_描述.png', '图02_描述.png'],
        image_after={0: 0, 1: 2},  # 图01放在第0段后，图02放在第2段后
        formula_after={1: 'eq1'},   # eq1公式放在第1段后
    ),
    # Module(
    #     title='第二个模块标题',
    #     summary='（2）第二个模块标题，...',
    #     description='...',
    #     images=[],
    #     image_after={},
    # ),
]


# ============================================================
# 【用户配置区】图注定义
# ============================================================
# key = 图片文件名，value = 图注文字（不含"图 N"前缀，脚本自动编号）
fig_captions = {
    # '整体框架.png': '整体研究框架',
    # '图01_描述.png': '数据采集实验装置',
    # '图02_描述.png': '实验结果对比',
}


# ============================================================
# 校验函数（一般不需要修改）
# ============================================================
def validate_modules():
    """校验模块定义的正确性：索引范围、公式残留、图片重复"""
    FORMULA_PATTERNS = [
        r'[A-Za-z_]+\s*=\s*[A-Za-z_]',
        r'[=+\-*/]\s*[A-Za-z_]+\s*[=+\-*/]',
        r'max\s*\(|min\s*\(|argmax|argmin',
        r'\|\|.*\|\|',
        r'[αβγλδεθΦΨ]\s*[=+\-*/^]',
    ]

    all_imgs = []
    for module in MODULES:
        para_count = len([t for t in module.description.split('\n') if t.strip()])
        for f_idx in module.formula_after:
            if f_idx >= para_count:
                raise ValueError(
                    f'模块"{module.title}"的formula_after索引{f_idx}超出范围'
                    f'（共{para_count}段，最大索引{para_count - 1}）'
                )
        for i_idx in module.image_after:
            if i_idx >= para_count:
                raise ValueError(
                    f'模块"{module.title}"的image_after索引{i_idx}超出范围'
                    f'（共{para_count}段，最大索引{para_count - 1}）'
                )
        for pattern in FORMULA_PATTERNS:
            matches = re.findall(pattern, module.description)
            if matches:
                print(f'警告: 模块"{module.title}"的description中可能包含公式表达式: {matches}')
        all_imgs.extend(module.images)

    # 检查图片重复
    from collections import Counter
    dupes = [k for k, v in Counter(all_imgs).items() if v > 1]
    if dupes:
        raise ValueError(f'图片重复使用: {dupes}')

    print('模块校验通过')


# ============================================================
# 章节写入函数（根据你的内容修改）
# ============================================================
def write_section1(cell):
    """
    一、思想品德与业务学习情况自述
    内容：思想政治、课程成绩、学习态度、学术道德、身体素质
    """
    clear_cell(cell)

    add_para(cell,
        '本人自XXXX年X月入学以来，严格遵守学校的各项规章制度，'
        '在思想上积极向党组织靠拢，认真学习马克思主义基本理论...'
    )
    add_para(cell,
        '在学习上，本人勤奋努力，积极向上...'
        '目前培养计划总学分XX分，现已圆满修满全部学分，总平均绩点达X.XX，无任何重修课程。'
    )
    add_para(cell,
        '结合XXX专业的培养要求，本人具备了扎实的学科知识...'
    )
    add_para(cell,
        '在学术道德方面，本人严格遵守学术规范和道德准则...'
    )
    add_para(cell,
        '在身体素质方面，本人平时积极参加体育锻炼，保持良好的身心状态...'
    )


def write_section2(cell):
    """
    二、已完成的科研工作（三层行文逻辑）

    第一层：开篇总述（研究目的 + 整体框架图 + 模块预告）
    第二层：N部分简要概述（每个模块1-2句话）
    第三层：详细展开（"已完成工作：" + 各模块详细描述）
    """
    clear_cell(cell)

    # ============================================================
    # 第一层：开篇总述
    # ============================================================
    add_para(cell,
        '本文的研究目的是针对XXX领域存在的XXX问题，设计一种XXX方法，'
        '以实现XXX目标。如图1所示为本文的整体研究框架。'
    )
    add_img(cell, img('整体框架.png'), width=Inches(5.5))
    add_caption(cell, '图 1 整体研究框架')

    # ============================================================
    # 第二层：N部分简要概述
    # ============================================================
    add_para(cell, '目前已完成的工作主要包括以下方面：')

    for module in MODULES:
        add_para(cell, module.summary)

    # ============================================================
    # 第三层：详细展开
    # ============================================================
    add_para(cell, '已完成工作：')

    fig_num = 2  # 图1已在概述中使用
    for mod_idx, module in enumerate(MODULES):
        # 模块标题，带编号
        add_para(cell, f'（{mod_idx + 1}）{module.title}', bold=True)

        paragraphs = [t for t in module.description.split('\n') if t.strip()]
        placed_imgs = set()
        placed_formulas = set()

        for para_idx, para_text in enumerate(paragraphs):
            add_para(cell, para_text)

            # 插入公式（紧跟描述段落）
            for f_idx, f_name in module.formula_after.items():
                if f_idx == para_idx and f_name not in placed_formulas:
                    add_formula(cell, f_name)
                    placed_formulas.add(f_name)

            # 插入图片（紧跟引用段落）
            for img_idx, target_para in module.image_after.items():
                if target_para == para_idx and img_idx not in placed_imgs:
                    add_img(cell, img(module.images[img_idx]), width=Inches(5.0))
                    caption = fig_captions.get(module.images[img_idx], module.title)
                    add_caption(cell, f'图 {fig_num} {caption}')
                    fig_num += 1
                    placed_imgs.add(img_idx)

        # 插入剩余未放置的图片
        for img_idx, img_name in enumerate(module.images):
            if img_idx not in placed_imgs:
                add_img(cell, img(img_name), width=Inches(5.0))
                caption = fig_captions.get(img_name, module.title)
                add_caption(cell, f'图 {fig_num} {caption}')
                fig_num += 1


def write_section3(cell):
    """
    三、下一步科研计划
    内容：分阶段科研安排（日期范围 + 详细描述）
    """
    clear_cell(cell)

    add_para(cell, '根据目前的研究进展和课题计划，下一步的科研工作主要包括以下阶段：')

    add_para(cell,
        'XXXX.XX-XXXX.XX：第一阶段计划...'
        '具体而言，将重点XXX...',
        first_indent='440'
    )
    add_para(cell,
        'XXXX.XX-XXXX.XX：第二阶段计划...'
        '具体而言，将重点XXX...',
        first_indent='440'
    )
    add_para(cell,
        'XXXX.XX-XXXX.XX：第三阶段计划...'
        '撰写学位论文，进行修改完善和答辩准备。',
        first_indent='440'
    )


# ============================================================
# 主函数（一般不需要修改）
# ============================================================
def main():
    validate_modules()

    # 从模板恢复
    shutil.copy2(TEMPLATE_FILE, WRITE_FILE)
    print('Restored from template')

    doc = Document(WRITE_FILE)
    assert len(doc.tables) >= 1, '模板文件中未找到表格'
    table = doc.tables[0]
    assert len(table.rows) >= 6, '模板表格至少需要6行（3个章节 × 2行）'

    print('Writing Section 1...')
    write_section1(table.rows[1].cells[0])

    print('Writing Section 2...')
    write_section2(table.rows[3].cells[0])

    print('Writing Section 3...')
    write_section3(table.rows[5].cells[0])

    doc.save(WRITE_FILE)
    print(f'Document saved: {WRITE_FILE}')

    # 统计
    total_text = ''
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                total_text += p.text
    char_count = len(total_text)
    print(f'总文字量: {char_count} 字')
    if char_count < 6000:
        print(f'警告: 文字量不足6000字，当前{char_count}字')
    else:
        print(f'文字量检查通过: {char_count} >= 6000')


if __name__ == '__main__':
    main()
