"""
中期考核报告生成脚本 - 示例
基于模板文档重写三个章节内容
使用 lyw-模板（工程实现型，三层行文逻辑）

运行: python rewrite_report.py
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
# 配置
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_FILE = os.path.join(BASE_DIR, '模板.docx')
WRITE_FILE = os.path.join(BASE_DIR, '写入文件.docx')
IMG_DIR = os.path.join(BASE_DIR, '图片')

# ============================================================
# 公式定义（LaTeX）
# ============================================================
FORMULAS: dict[str, str] = {
    'nmf': r'V = WH + E',
    'vaf': r'\text{VAF} = 1 - \frac{\| V - WH \|_F^2}{\| V \|_F^2}',
    'loss': r'L_{total} = \alpha \| \hat{h}_S - h_{GT} \|^2 + (1 - \alpha) \| \hat{h}_S - \hat{h}_T \|^2',
    'finetune': r'\theta_{final} = \theta_{pre} - \eta \nabla_\theta L_{hard}(\theta_{pre}, D_{calib})',
    'hdiff': r'H_{diff} = \max(H_{healthy} - H_{patient},\ 0)',
    'copt': r'C_{opt} = \alpha \cdot C_{ff} + (1 - \alpha) \cdot C_{fb}',
}

# ============================================================
# 图片路径
# ============================================================
def img(name: str) -> str:
    path = os.path.join(IMG_DIR, name)
    if not os.path.exists(path):
        raise FileNotFoundError(f'图片不存在: {path}')
    return path

# ============================================================
# 工具函数
# ============================================================
def set_run_font(run, font_name='Times New Roman', font_size=Pt(12), bold=None, east_asia='宋体'):
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
    for i in range(len(cell.paragraphs) - 1, 0, -1):
        p = cell.paragraphs[i]
        p._element.getparent().remove(p._element)
    first_p = cell.paragraphs[0]
    for run in first_p.runs:
        run.text = ''
    for r in first_p._element.findall(qn('w:r')):
        first_p._element.remove(r)


def _apply_spacing(pPr):
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = etree.SubElement(pPr, qn('w:spacing'))
    spacing.set(qn('w:line'), '360')
    spacing.set(qn('w:lineRule'), 'auto')
    spacing.set(qn('w:after'), '0')
    return spacing


def _apply_indent(pPr, first_indent='480', left_indent='34', first_line_chars='200'):
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
    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(img_path, width=width)
    pPr = p._element.get_or_add_pPr()
    _apply_spacing(pPr)
    return p


def add_caption(cell, text, font_size=Pt(10)):
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
    latex = FORMULAS.get(formula_name)
    if latex is None:
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
        # 先用pdflatex编译确保LaTeX公式可渲染
        subprocess.run(
            ['pdflatex', '-interaction=nonstopmode', tex_path],
            capture_output=True, timeout=30, cwd=tmp_dir
        )
        # 再用pandoc转换为docx
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
    title: str
    summary: str
    description: str
    images: list[str]
    image_after: dict[int, int] = field(default_factory=dict)
    formula_after: dict[int, str] = field(default_factory=dict)


# ============================================================
# 模块定义
# ============================================================
MODULES = [
    Module(
        title='多模态数据采集与预处理',
        summary='（1）多模态数据采集与预处理，搭建多模态数据采集平台，采集sEMG和关节运动学数据，'
                '并进行信号滤波、包络提取和标准化等预处理操作，为后续协同建模提供高质量数据基础。',
        description=(
            '为了建立准确的肌肉协同模型，本人搭建了多模态数据采集平台。'
            '采集系统集成了Delsys Trigno无线表面肌电传感器（采样率1926 Hz）和Biometrics电子角度传感器（采样率1000 Hz），'
            '通过硬件触发实现双模态数据流的严格同步。实验装置采用定制阻抗滑轨平台，约束上肢运动严格限制在水平面内，'
            '有效增强伸肌群激活，确保高保真的神经肌肉表征记录。\n'
            '为实现上肢运动在水平面内的精确神经肌肉表征，共采集了12块关键浅层肌群的信号。'
            '腕关节组包括伸肌群（桡侧腕长伸肌ECRL、指伸肌ED、尺侧腕伸肌ECU）和屈肌群（尺侧腕屈肌FCU、指浅屈肌FDS、桡侧腕屈肌FCR）。'
            '肘关节组包括屈肌群（肱桡肌BR、肱二头肌长头BBLH、肱肌BRA、肱二头肌短头BBSH）和伸肌群（肱三头肌长头TBLH、肱三头肌外侧头TBLATH）。'
            '每位受试者每个关节连续运动下每组采集时长为42秒，每组采集结束休息3分钟后进行下一组，共采集三组。\n'
            '数据预处理方面，首先对原始sEMG信号进行带通滤波（20-450 Hz）和陷波滤波（50 Hz工频干扰消除），'
            '然后通过全波整流和低通滤波提取信号包络，最后进行Z-score标准化处理。'
            '关节角度信号首先进行上采样以保证与肌电包络的样本点数目一致，随后计算角速度并进行中值滤波和平均值滤波平滑处理。\n'
            '如图2所示为VAF（方差贡献率）与协同数的关系分析。'
            'VAF的定义为：\n'
            '在协同数r从1到6的范围内，VAF呈非线性增长趋势。'
            '肘关节在r等于3时VAF达到92.4%，腕关节在r等于3时VAF达到92.6%，均超过90%的重构精度阈值。'
            '因此选取r等于3作为表征平面运动的最小协同数目。该高压缩比定量支持了中枢神经系统采用模块化组织策略来简化多关节控制的假说。\n'
            '不同阻力和疲劳因素对肌肉协同的影响分析表明，不同阻力下肌肉协同结构较为稳健，但肌肉疲劳会导致协同稳定性降低。'
            '该分析为后续协同模型建立提供了重要依据，且确定了稳定的肌肉协同提取方法。'
        ),
        images=['图04_多模态数据采集装置.jpeg', '图05_数据预处理.jpeg', '协同模式提取.png'],
        image_after={0: 0, 1: 1, 2: 3},
        formula_after={3: 'vaf'},
    ),
    Module(
        title='跨模态泛化协同建模',
        summary='（2）跨模态泛化协同建模，采用NMF提取肌肉协同特征，设计ResMamba混合骨干网络，'
                '通过非对称跨模态知识蒸馏和迁移学习，构建从运动学到肌肉协同空间的鲁棒映射模型。',
        description=(
            '为了从非侵入性运动学信号中鲁棒地重建肌肉协同模式，本人设计了一种基于ResMamba混合架构与跨模态知识蒸馏的协同预测框架。'
            '该框架首先通过非负矩阵分解从表面肌电信号中提取肌肉协同特征。'
            '该分解过程表示为：\n'
            '空间协同矩阵反映运动基元的肌肉权重分布，时间激活系数反映各协同的时间激活模式，残差项衡量重构误差。'
            '通过乘法更新规则求解，W矩阵列归一化以消除幅值缩放歧义。所提方法在健康场景下腕关节和肘关节任务的平均预测相关系数分别达到0.857和0.808。\n'
            'ResMamba混合骨干网络串联残差卷积和Mamba模块。残差卷积部分通过两个级联残差块进行局部时空特征提取，第一个残差块采用步长为2进行下采样，扩大感受野并降低计算量；'
            '第二个残差块保持时间分辨率以提取高频运动学特征。Mamba模块利用状态空间模型的线性复杂度和长程依赖建模能力，高效捕捉非平稳上肢运动中的长时间序列依赖关系。\n'
            '非对称跨模态知识蒸馏机制通过教师-学生框架，将表面肌电模态的神经控制先验迁移到纯运动学学生网络。'
            '教师网络以运动学数据和中值频率特征作为输入，学生网络仅以运动学数据作为输入。'
            '复合损失函数定义为：\n'
            '其中第一项为硬损失，衡量学生预测与真实标签之间的均方误差；第二项为蒸馏损失，量化学生预测与教师软目标之间的差异。'
            '软目标编码了多模态特征空间中的额外监督信息，包括时间平滑的激活趋势和非线性协调模式。'
            '权重系数alpha等于0.5，通过实验经验确定。\n'
            '迁移学习采用两阶段策略以处理被试间变异性。'
            '第一阶段在多个源被试数据上进行群体级预训练，构建鲁棒的特征表示。'
            '第二阶段为目标被试进行个性化微调，采用留一被试交叉验证范式。'
            '如图6所示为最优微调比例的探寻结果，在微调比例为15%处存在明显的膝点，'
            '该点之前为快速生长区，精度提升显著；该点之后为收益递减区，精度增益趋于平缓。'
            '因此选取前15%的校准数据进行个性化参数适配。\n'
            '如图7所示为所提方法在腕关节和肘关节数据集上与其他方法的对比结果。'
            '所提方法S9在所有测试条件下均取得最高的重建精度，腕关节任务平均相关系数达到0.857，肘关节任务达到0.808。'
            '与单流基线网络相比，S9实现了约25%至33%的性能提升。'
            '与混合架构基线相比，S9保持约15%的显著领先优势。\n'
            '如图8所示为协同激活系数的时序重建效果。'
            '所提方法在腕关节三个协同上分别达到0.909、0.925和0.891的相关系数，'
            '肘关节三个协同上分别达到0.873、0.847和0.899的相关系数，验证了高保真协同重建能力。'
        ),
        images=['图07_蒸馏架构与ResNet_Mamba.jpeg', '图08_最佳微调比例探寻.png',
                '图09_迁移微调策略.jpeg', '图10_模型在腕关节和肘关节数据集上的性能指标.png',
                '图12_高保真协同时序重建效能.png'],
        image_after={0: 0, 1: 5, 2: 5, 3: 6, 4: 7},
        formula_after={0: 'nmf', 3: 'loss', 6: 'finetune'},
    ),
    Module(
        title='实时协同缺损评估与前馈按需辅助解码',
        summary='（3）实时协同缺损评估与前馈按需辅助解码，基于Assist-as-Needed机制，'
                '实时评估患者协同缺损程度并生成前馈补偿信号，实现个性化辅助控制。',
        description=(
            '为了实现个性化的按需辅助控制，本人设计了基于Assist-as-Needed机制的实时协同缺损评估与前馈解码模块。'
            '该模块的核心思想是仅在患者需要时提供最小必要辅助量，避免过度刺激导致的肌肉疲劳和废用性萎缩。\n'
            '协同缺损向量定义为健康协同模式与患者残余协同之差的非负截断：\n'
            '其中H healthy为离线建立的健康协同先验，H patient为患者当前的残余协同估计，负值截断确保辅助量非负。'
            '该向量量化了患者所需的最小辅助量，反映了运动执行过程中神经肌肉控制能力的缺失程度。\n'
            '离线健康协同先验与实时协同缺损评估结合，生成前馈补偿信号。'
            '该信号与MISO-MFAC自适应控制器的反馈信号融合，形成动态控制律，最终驱动FES执行端产生电刺激。'
            '系统实时采集关节角度反馈，形成完整的闭环控制回路。该策略在运动初期协同缺损较大时强化前馈补偿，'
            '在运动后期协同恢复时抑制刺激强度，从而在保证运动质量的同时减轻肌肉疲劳。'
        ),
        images=[],
        image_after={},
        formula_after={1: 'hdiff'},
    ),
    Module(
        title='基于MFAC的融合闭环控制律构建',
        summary='（4）基于MFAC的融合闭环控制律构建，将前馈补偿信号与MISO-MFAC反馈控制信号融合，'
                '通过数据驱动线性化和自适应权重调整，实现抗疲劳的闭环轨迹跟踪控制。',
        description=(
            '为了将肌肉协同模型与控制算法有机结合，本人设计了基于无模型自适应控制的融合闭环控制律。'
            '该控制律将前馈补偿信号与多输入单输出无模型自适应控制器的反馈信号进行动态融合。'
            '最优控制律定义为：\n'
            '其中C ff为基于肌肉协同模型的前馈补偿信号，C fb为MISO-MFAC控制器的反馈修正信号，alpha为动态权重系数。'
            '该融合策略利用了前馈信号的快速响应特性和反馈信号的误差修正能力，在运动初期协同缺损较大时强化前馈补偿，'
            '在运动后期协同恢复时增大反馈权重以抑制刺激强度，从而在保证轨迹跟踪精度的同时减轻肌肉疲劳。\n'
            '无模型自适应控制器采用数据驱动线性化方法，无需建立被控对象的精确数学模型。'
            '控制器通过在线伪梯度估实现对时变系统的自适应跟踪，利用紧格式动态线性化技术将非线性系统在每个工作点处等价为时变线性系统。'
            '控制器的参数更新律根据跟踪误差动态调整控制增益，保证闭环系统的收敛性和稳定性。\n'
            '如图10所示为MFAC融合闭环控制律的构建过程。'
            '离线健康协同先验与实时协同缺损评估结合生成前馈补偿信号，与MISO-MFAC自适应控制器的反馈信号融合形成动态控制律，'
            '最终驱动FES执行端产生电刺激。'
        ),
        images=['基于MFAC的融合闭环控制律构建.png'],
        image_after={0: 3},
        formula_after={0: 'copt'},
    ),
    Module(
        title='实验平台搭建',
        summary='（5）实验平台搭建，完成上位机中枢控制、多通道FES刺激器和角度传感器的系统集成，'
                '搭建具备人体闭环验证条件的实验平台。',
        description=(
            '为了验证所提方法的可行性和有效性，本人搭建了完整的上肢康复FES实验平台。'
            '如图11所示为康复系统的整体框图。系统由上位机控制中枢、多通道电刺激器、电源模块、可穿戴式电极、'
            '角度传感器和蓝牙接收器组成。上位机运行控制算法，通过蓝牙与电刺激器通信，发送刺激参数指令。'
            '角度传感器实时采集关节角度反馈，形成完整的闭环控制回路。\n'
            '如图12所示为上肢康复系统的实际实验平台。'
            '上位机采用笔记本电脑运行MATLAB控制程序，通过串口与多通道电刺激器通信。'
            '多通道电刺激器包含MCU控制单元和多路恒流源模块，能够独立控制多块肌肉的刺激强度和脉冲宽度。'
            '电源模块将12V直流电源转换为5V和90V，分别供给控制电路和刺激输出电路。'
            '实验中在受试者上臂和前臂分别放置电极阵列，覆盖肱二头肌、肱三头肌等目标肌群。'
            '角度传感器固定于关节旋转轴处，通过蓝牙接收器将角度数据实时传输至上位机，实现闭环反馈控制。\n'
            '如图13所示为闭环控制效能评估框架。'
            '实验范式设计包括腕关节和肘关节的连续重复性运动任务。'
            '量化评估指标涵盖轨迹跟踪精度、迭代间收敛速度和输出平滑度三个维度。'
            '核心验证目标包括克服固有非线性死区、抑制时变增益衰减以及验证前馈反馈融合策略的有效性。'
            '鲁棒性验证通过设置多周期循环任务，利用运动中自然产生的肌肉疲劳效应，评估控制律的动态补偿能力。'
        ),
        images=['图13_康复系统框图.jpeg', '图14_上肢康复系统实验平台.jpeg', '闭环控制效能评估.png'],
        image_after={0: 0, 1: 1, 2: 2},
    ),
    Module(
        title='学术成果积累',
        summary='（6）学术成果积累，在跨模态协同解码和上肢康复控制领域取得多项研究成果。',
        description=(
            '在跨模态协同解码和上肢康复控制领域，本人已取得多项研究成果。'
            '目前已完成的学术成果包括：\n'
            '论文1：作者1, 作者2, 作者3, et al. '
            '"论文标题" '
            '—— 期刊名称，状态。\n'
            '论文2：作者1, 作者2, et al. '
            '"论文标题" '
            '—— 期刊名称, 年份, 卷: 页码，已发表。\n'
            '论文3：作者1, 作者2, et al. '
            '"论文标题" '
            '—— 会议名称, 年份: 页码，已发表。\n'
            '软著：软件著作权名称 —— 已授权。\n'
            '实用新型专利：专利名称 —— 已授权。\n'
            '比赛：竞赛名称及获奖等级。\n'
            '综上所述，本人已完成了从数据采集、跨模态建模到闭环控制的完整技术链路搭建，'
            '形成了具有自主知识产权的上肢康复FES控制系统。'
        ),
        images=[],
        image_after={},
    ),
]

# ============================================================
# 校验函数
# ============================================================
def validate_modules():
    """校验模块定义的正确性"""
    FORMULA_PATTERNS = [
        r'[A-Za-z_]+\s*=\s*[A-Za-z_]',
        r'[=+\-*/]\s*[A-Za-z_]+\s*[=+\-*/]',
        r'max\s*\(|min\s*\(|argmax|argmin',
        r'\|\|.*\|\|',
        r'[αβγλδεθΦΨ]\s*[=+\-*/^]',
    ]
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
    print('模块校验通过')


# ============================================================
# 章节写入
# ============================================================
def write_section1(cell):
    """一、思想品德与业务学习情况自述"""
    clear_cell(cell)

    add_para(cell,
        '本人自XXXX年X月入学以来，严格遵守学校的各项规章制度，在思想上积极向党组织靠拢，'
        '认真学习马克思主义基本理论，拥护党的路线方针政策，坚持四项基本原则。'
        '在学习上，本人勤奋努力，积极向上，不断追求个人的成长和进步。'
        '本人学习态度端正、踏实认真，学习成绩优良。目前培养计划总学分XX分，'
        '现已圆满修满全部学分，总平均绩点达X.XX，无任何重修课程。'
    )
    add_para(cell,
        '结合XXX专业的培养要求，本人具备了扎实的学科知识。'
        '通过系统的课程学习，本人在《应用非线性控制》、《线性系统理论》、《数字信号处理》'
        '以及《矩阵分析引论》等核心基础课程，以及《机器学习》、《智能机器人控制技术》等前沿课程中均取得了优异成绩，'
        '构建了系统且完备的专业理论框架。在科研上，本人认真钻研，积极探索，注重提升自己的专业技能，'
        '努力跟进行业的最新动态，不断学习新的工具和技术，以提高工作效率。'
        '本人能够将理论知识与实践相互结合，进而实现理论知识的实际应用和创新。'
        '同时，本人具备良好的研究、分析并解决问题的能力，能够进行有效的文献调研、设计研究方案、收集和处理数据，'
        '并得出准确、可靠的结论。'
    )
    add_para(cell,
        '在学术道德方面，本人严格遵守学术规范和道德准则，在《工程伦理学》等课程的指导下，'
        '本人深知尊重知识产权的重要性，坚决杜绝抄袭他人的作品或研究成果。'
        '在日常科研中，本人能够正确引用和参考他人的工作，并妥善处理合作研究中的学术诚信问题，'
        '始终保持严谨求实的治学态度。在身体素质方面，本人平时积极参加体育锻炼，'
        '保持良好的身心状态，为科研工作提供了良好的体能基础。'
    )


def write_section2(cell):
    """二、已完成的科研工作（三层行文逻辑）"""
    clear_cell(cell)

    # ============================================================
    # 第一层：开篇总述
    # ============================================================
    add_para(cell,
        '本文的研究目的是针对上肢多关节复杂肌肉控制存在的控制信号冗余和控制器设计复杂问题，'
        '设计一种基于表面肌电信号和运动信息的肌肉协同模型的迭代学习控制算法，'
        '以实现基于肌肉协同理论的功能性电刺激闭环控制，'
        '解决复杂肌肉电刺激中多控制信号的参数调制问题，实现多肌肉多控制信号的降维控制。'
        '如图1所示为本文的整体研究框架。'
    )
    add_img(cell, img('整体框架.png'), width=Inches(5.5))
    add_caption(cell, '图 1 整体研究框架')

    # ============================================================
    # 第二层：六部分简要概述
    # ============================================================
    add_para(cell, '目前已完成的工作主要包括以下六个方面：')

    for module in MODULES:
        add_para(cell, module.summary)

    # ============================================================
    # 第三层：详细展开
    # ============================================================
    add_para(cell, '已完成工作：')

    fig_num = 2  # 图1已在概述中使用
    fig_captions = {
        '整体框架.png': '整体研究框架',
        '图04_多模态数据采集装置.jpeg': '数据采集实验装置',
        '图05_数据预处理.jpeg': '数据预处理流程',
        '协同模式提取.png': 'VAF与协同数关系',
        '图07_蒸馏架构与ResNet_Mamba.jpeg': '蒸馏架构与ResNet-Mamba骨干网络',
        '图08_最佳微调比例探寻.png': '最佳微调比例探寻',
        '图09_迁移微调策略.jpeg': '迁移学习框架',
        '图10_模型在腕关节和肘关节数据集上的性能指标.png': '模型性能对比',
        '图12_高保真协同时序重建效能.png': '协同时序重建效能',
        '基于MFAC的融合闭环控制律构建.png': 'MFAC融合闭环控制律',
        '图13_康复系统框图.jpeg': '康复系统框图',
        '图14_上肢康复系统实验平台.jpeg': '上肢康复系统实验平台',
        '闭环控制效能评估.png': '闭环控制效能评估框架',
    }
    for mod_idx, module in enumerate(MODULES):
        # 模块标题，带（1）（2）...编号，与概述对应
        add_para(cell, f'（{mod_idx + 1}）{module.title}', bold=True)
        paragraphs = [t for t in module.description.split('\n') if t.strip()]
        placed_imgs = set()
        placed_formulas = set()

        for para_idx, para_text in enumerate(paragraphs):
            add_para(cell, para_text)

            # 插入公式
            for f_idx, f_name in module.formula_after.items():
                if f_idx == para_idx and f_name not in placed_formulas:
                    add_formula(cell, f_name)
                    placed_formulas.add(f_name)

            # 插入图片
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
    """三、下一步科研计划"""
    clear_cell(cell)

    add_para(cell, '根据目前的研究进展和课题计划，下一步的科研工作主要包括以下三个阶段：')

    add_para(cell,
        'XXXX.XX-XXXX.XX：针对物理病态时变特征，进一步优化前馈导引与反馈修正的动态融合机制。'
        '具体而言，将重点完善MFAC控制器在肌肉疲劳场景下的鲁棒性，通过引入自适应权重调度策略，'
        '使控制器能够在肌肉疲劳导致的增益衰减条件下保持稳定的轨迹跟踪性能。'
        '同时，优化协同缺损评估算法的实时性，降低计算延迟以满足临床实时控制的需求。',
        first_indent='440'
    )
    add_para(cell,
        'XXXX.XX-XXXX.XX：扩展实验验证的广度和深度。'
        '对更多健康受试者施加本文设计的功能性电刺激控制方法，验证所提出方法在多通道FES降维控制上的可行性和泛化能力。'
        '同时，收集更多卒中患者的临床数据，评估所提方法在病理条件下的适应性和康复效果。'
        '在此基础上，优化系统的工程化设计，提高实验平台的稳定性和易用性。',
        first_indent='440'
    )
    add_para(cell,
        'XXXX.XX-XXXX.XX：撰写大论文，对整体工作做详细总结。'
        '系统梳理从数据采集、跨模态协同建模、协同缺损评估到闭环控制的完整技术链路，'
        '形成完整的学位论文，并进行修改完善和答辩准备。',
        first_indent='440'
    )


# ============================================================
# 主函数
# ============================================================
def main():
    validate_modules()

    # 从模板恢复
    shutil.copy2(TEMPLATE_FILE, WRITE_FILE)
    print('Restored from template')

    doc = Document(WRITE_FILE)
    assert len(doc.tables) >= 1
    table = doc.tables[0]
    assert len(table.rows) >= 6

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
