"""
中期考核报告生成脚本模板 v3.1
根据实际工作目录和材料进行修改

使用方法：
1. 修改 BASE_DIR 为实际工作目录
2. 图片来源：优先从PPT/论文提取，也可使用图片/目录
3. 追溯每张图片来源（PPT/论文），确认内容与图注对应
4. 修改各章节内容（write_section1/2/3函数）
5. 运行: E:/Anaconda/python.exe rewrite_report.py

注意：
- 不要使用 conda run -n base python -c "多行代码"（Windows不支持）
- 不要信任图片文件名，必须用视觉模型核验图片实际内容
- 公式使用 pandoc LaTeX → OMML 转换，确保渲染正确（需要安装 pandoc）
- 不再需要写入文件(3).docx、模板.docx、中期考核相关资料.docx

环境要求：
- Python >= 3.7
- 依赖：python-docx, PyMuPDF (fitz), Pillow, lxml
"""

from __future__ import annotations

import os
import shutil
import sys
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
        图片文件的完整路径，如果未找到则返回 None

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
# 章节写入函数 - 根据实际内容修改
# ============================================================
def write_section1(cell) -> None:
    """一、思想品德与业务学习情况自述

    内容要点（根据研究生成绩单填写）：
    - 思想政治：政治态度、理论学习、制度遵守
    - 课程成绩：总学分、绩点、核心课程及分数
    - 学习态度：勤奋努力、积极向上
    - 科研能力：参与项目、技能掌握
    - 学术道德：规范遵守、课程支撑
    - 身体素质：锻炼习惯、身心健康
    """
    clear_cell(cell)

    add_para(cell, '思想政治方面内容...')
    add_para(cell, '课程成绩方面内容（根据成绩单填写）...')
    add_para(cell, '学习态度方面内容...')
    add_para(cell, '科研能力方面内容...')
    add_para(cell, '学术道德方面内容...')
    add_para(cell, '身体素质方面内容...')


def write_section2(cell) -> None:
    """二、已完成的科研工作

    三层行文逻辑（必须遵循）：
    1. 开篇总述：研究目的 + 整体框架 + 框架图
    2. 六部分简要概述：每个模块1-2句话概括
    3. 详细展开："已完成工作：" + 每个模块详细描述

    每个模块展开结构：
    - 背景与目的（为什么要做）
    - 方法描述（怎么做）
    - 图片展示（配图说明）
    - 实验验证（数据支撑）
    - 创新点/结论
    """
    clear_cell(cell)

    # ============================================================
    # 第一层：开篇总述
    # ============================================================
    add_para(cell, '本文的研究目的是...如图1所示为本文的整体研究框架...')
    add_img(cell, img('图1.png'), width=Inches(5.5))
    add_caption(cell, '图1 整体研究框架')

    # ============================================================
    # 第二层：六部分简要概述
    # ============================================================
    add_para(cell, '目前已完成的工作主要包括以下六个方面：')

    add_para(cell,
        '（1）多模态数据采集与预处理，搭建多模态数据采集平台，采集sEMG和关节运动学数据，'
        '并进行信号滤波、包络提取和标准化等预处理操作，为后续协同建模提供高质量数据基础。')

    add_para(cell,
        '（2）跨模态泛化协同建模，采用NMF提取肌肉协同特征，设计ResMamba混合骨干网络，'
        '通过非对称跨模态知识蒸馏和迁移学习，构建从运动学到肌肉协同空间的鲁棒映射模型。')

    add_para(cell,
        '（3）实时协同缺损评估与前馈按需辅助解码，基于Assist-as-Needed机制，'
        '实时评估患者协同缺损程度并生成前馈补偿信号，实现个性化辅助控制。')

    add_para(cell,
        '（4）基于MFAC的融合闭环控制律构建，将前馈补偿信号与MISO-MFAC反馈控制信号融合，'
        '通过数据驱动线性化和自适应权重调整，实现抗疲劳的闭环轨迹跟踪控制。')

    add_para(cell,
        '（5）实验平台搭建，完成上位机中枢控制、多通道FES刺激器和角度传感器的系统集成，'
        '搭建具备人体闭环验证条件的实验平台。')

    add_para(cell,
        '（6）学术成果积累，在跨模态协同解码和上肢康复控制领域取得多项研究成果。')

    # ============================================================
    # 第三层：详细展开
    # ============================================================
    add_para(cell, '已完成工作：')

    # 模块1：多模态数据采集与预处理
    add_para(cell, '（1）多模态数据采集与预处理。详细描述...')
    add_img(cell, img('图2.png'), width=Inches(4.5))
    add_caption(cell, '图2 数据采集实验装置')
    # ... 更多内容 ...

    # 模块2：跨模态协同建模
    add_para(cell, '（2）跨模态泛化协同建模。详细描述...')
    # ... 包含公式 ...
    # from omml_formulas import create_formula_paragraph
    # create_formula_paragraph(cell, 'loss')
    add_img(cell, img('图6.png'), width=Inches(5.0))
    add_caption(cell, '图6 跨模态蒸馏架构与ResNet-Mamba骨干网络')

    # 模块3-5 同理...

    # 模块6：学术成果
    add_para(cell, '（6）已取得的学术成果。如图16所示为目前已取得的各项学术成果。')
    add_img(cell, img('图16.png'), width=Inches(5.0))
    add_caption(cell, '图16 已取得学术成果')
    add_para(cell, '论文1：...')
    add_para(cell, '论文2：...')
    add_para(cell, '软著：...')
    add_para(cell, '专利：...')
    add_para(cell, '比赛：...')

    add_para(cell,
        '综上所述，本人已完成了从数据采集、跨模态建模到闭环控制的完整技术链路搭建，'
        '形成了具有自主知识产权的上肢康复FES控制系统。')


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
    add_para(cell, '2026.07-2026.08：第一阶段内容...', first_indent='440')
    add_para(cell, '2026.09-2026.11：第二阶段内容...', first_indent='440')
    add_para(cell, '2026.12-2027.04：第三阶段内容...', first_indent='440')


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
