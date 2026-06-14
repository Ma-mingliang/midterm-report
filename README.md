# 中期考核报告自动生成工具

基于 `python-docx` 的研究生中期考核报告自动生成工具。通过分析 PPT、论文 PDF 和成绩单，自动生成符合学校格式要求的中期考核报告 Word 文档。

## 功能特性

- **动态模块检测**：从 PPT 和论文中自动识别研究模块结构，不依赖预设模板
- **双模板支持**：lyw-模板（工程实现型）和 yl-模板（理论推导型）
- **OMML 公式渲染**：通过 pdflatex + pandoc 将 LaTeX 公式转换为 Word 原生公式编辑器格式
- **图片交叉验证**：视觉模型 + PPT 文字 + 论文描述三方验证，确保图注与内容一致
- **内联图片插入**：图片紧跟"如图X所示"引用位置，而非集中堆放在模块末尾
- **精确排版控制**：宋体 + Times New Roman 混合字体、1.5 倍行距、首行缩进 2 字符
- **三层行文逻辑**（lyw-模板）：开篇总述 → 六部分概述 → 详细展开
- **自动校验**：公式残留扫描、索引越界检查、图片数量与编号连续性验证

## 环境要求

### Python 依赖

```
python-docx
PyMuPDF (fitz)
Pillow
lxml
```

### 系统依赖

- **pandoc**：LaTeX 到 OMML 的公式转换
- **pdflatex**：LaTeX 公式编译（MiKTeX 或 TeX Live）

```bash
# 安装 pandoc
conda install -c conda-forge pandoc
# 或从 https://pandoc.org/installing.html 下载

# 安装 MiKTeX（含 pdflatex）
# 从 https://miktex.org/download 下载安装
```

### Python 环境

推荐使用 Anaconda：

```bash
conda create -n midterm python=3.11
conda activate midterm
pip install python-docx PyMuPDF Pillow lxml
```

## 快速开始

### 1. 准备文件

在工作目录下放置以下文件：

```
工作目录/
├── 模板.docx              # 格式模板（封面页、表格结构）
├── 研究生成绩单.pdf        # 课程成绩、学分、绩点
├── 论文.pdf               # 技术内容来源
├── 中期ppt.pdf            # 工作框架来源
├── 中期考核相关资料.docx   # 考核要求
├── 图片/                   # 所有待插入报告的图片
│   ├── 图01_xxx.png
│   ├── 图02_xxx.jpg
│   └── ...
└── rewrite_report.py       # 生成脚本
```

### 2. 修改脚本配置

```python
BASE_DIR = r'你的工作目录路径'
TEMPLATE_FILE = os.path.join(BASE_DIR, '你的模板文件名.docx')
IMG_DIR = os.path.join(BASE_DIR, '图片')
```

### 3. 运行

```bash
python rewrite_report.py
# 或指定 Python 路径
E:/Anaconda/python.exe rewrite_report.py
```

### 4. 输出

生成 `写入文件.docx`，包含：
- 第一章：思想品德与业务学习情况自述
- 第二章：已完成的科研工作（三层行文逻辑）
- 第三章：下一步科研计划

## 详细使用流程

本工具遵循 6 阶段工作流：

### 阶段 1：需求收集与材料分析

1. **扫描材料清单**：确认成绩单、论文、PPT、图片目录齐全
2. **读取成绩单**：提取总学分、绩点、核心课程及分数
3. **分析 PPT 结构**：提取每页文字，识别页面类型（封面/目录/内容/总结），检测模块边界
4. **确认信息来源**：向用户确认检测到的模块结构是否正确

### 阶段 2：图片深度理解与映射

**这是整个流程中最关键的阶段。**

对每张图片执行：

1. **视觉模型查看**：描述图片类型、文字标注、具体内容
2. **PPT 上下文匹配**：在 PPT 中搜索相关页面，提取完整文字
3. **论文章节匹配**：找到论文中对应的章节和图注
4. **三方交叉验证**：视觉模型 ↔ PPT 文字 ↔ 论文描述，三者必须一致
5. **建立映射表**：记录文件名、含义、分配模块、图注、验证状态

**图片目录硬性要求：**
- 每张图片都必须在报告中被使用，不允许遗漏
- 每张图片只能被使用 1 次，不允许重复
- 图片总数即为报告中的图片总数

### 阶段 3：内容规划与行文逻辑

#### lyw-模板（工程实现型）

| 章节 | 内容 | 篇幅占比 |
|------|------|---------|
| 一、思想品德 | 思想政治、课程成绩、学习态度、身体素质 | ~10% |
| 二、已完成工作 | 整体框架 + N 个技术模块 | **~75%** |
| 三、科研计划 | 分阶段科研安排 | ~15% |

第二章三层行文逻辑：

```
第一层：开篇总述
  ├─ 研究目的（1-2句）
  ├─ 整体框架描述（配框架图）
  └─ "目前的已完成工作主要包括以下N个方面："

第二层：N部分简要概述
  ├─ 每个模块用1-2句话概括
  └─ 编号：（1）（2）（3）...

第三层：详细展开
  ├─ "已完成工作：" 过渡标记
  └─ 每个模块独立展开（方法、公式、图片、实验、创新点）
```

#### yl-模板（理论推导型）

- 无开篇总述和简要概述
- 直接"已完成工作："后进入各模块
- 模块用 `1. 2.` 编号，内部用 `1.1, 1.2` 小节
- 大量数学公式和推导

### 阶段 4：脚本生成

#### Module 数据结构

```python
@dataclass
class Module:
    title: str              # 模块标题
    summary: str            # 概述层的1-2句话概括
    description: str        # 详细展开的正文（用\n分隔段落）
    images: list[str]       # 图片文件名列表
    image_after: dict[int, int]    # {img_index: para_index} 图片插入位置
    formula_after: dict[int, str]  # {para_index: formula_name} 公式插入位置
```

#### 公式定义

```python
FORMULAS: dict[str, str] = {
    'nmf': r'V = WH + E',
    'vaf': r'\text{VAF} = 1 - \frac{\| V - WH \|_F^2}{\| V \|_F^2}',
    'loss': r'L_{total} = \alpha \| \hat{h}_S - h_{GT} \|^2 + (1 - \alpha) \| \hat{h}_S - \hat{h}_T \|^2',
}
```

公式通过 `add_formula(cell, 'loss')` 渲染为 OMML 格式，流程：

```
LaTeX 字符串 → pdflatex 编译 → pandoc 转换为 docx → 提取 oMath 元素 → 插入 Word 单元格
```

#### 内联插入逻辑

```python
paragraphs = module.description.split('\n')
for para_idx, para_text in enumerate(paragraphs):
    add_para(cell, para_text)
    # 先插入公式（紧跟描述段落）
    for f_idx, f_name in module.formula_after.items():
        if f_idx == para_idx:
            add_formula(cell, f_name)
    # 再插入图片（紧跟引用段落）
    for img_idx, target_para in module.image_after.items():
        if target_para == para_idx:
            add_img(cell, img(module.images[img_idx]))
            add_caption(cell, f'图 {fig_num} {caption}')
```

### 阶段 5：执行与验证

```bash
python rewrite_report.py
```

输出示例：

```
模块校验通过
Restored from template
Writing Section 1...
Writing Section 2...
Writing Section 3...
Document saved: 工作目录/写入文件.docx
总文字量: 6090 字
文字量检查通过: 6090 >= 6000
```

### 阶段 6：完整性核验

自动校验：
- OMML 公式数量 > 0
- 图片数量 = 图片目录中的全部图片
- 图编号连续无跳号
- 文字量 ≥ 6000 字
- formula_after/image_after 索引不越界
- description 中无公式表达式残留

## 格式规范

### 字体

| 元素 | 中文字体 | 英文字体 | 字号 |
|------|---------|---------|------|
| 章节标题 | 仿宋_GB2312 | — | 15pt |
| 正文段落 | 宋体 | Times New Roman | 12pt（小四） |
| 公式段落 | 宋体 | Times New Roman | 12pt |
| 图注 | 黑体 | Times New Roman | 10pt（五号） |

### 段落格式

| 属性 | 值 |
|------|-----|
| 行距 | 1.5 倍（line=360 twips, lineRule=auto） |
| 段后间距 | 0 |
| 正文首行缩进 | 480 twips + firstLineChars=200 |
| 正文左缩进 | 34 twips |
| 图注首行缩进 | 400 twips |

### 对齐方式

| 元素 | 对齐 |
|------|------|
| 正文段落 | 左对齐（jc=left） |
| 公式段落 | 居中（jc=center） |
| 图注段落 | 居中（jc=center） |

### 编号格式

- 小节编号：`（1）（2）（3）...`
- 图片编号：`图 1 图 2 图 3...`（连续无跳号）

## 模板选择指南

### lyw-模板（工程实现型）

**适用场景：**
- 工程类专业学位
- 偏实验验证的研究
- 图片数量多（≥15 张）
- 核心贡献是"做了一个系统/平台/原型"

**特点：**
- 三层行文逻辑（总述→概述→详述）
- 课程成绩逐门列出
- 统计分析丰富（ANOVA、Tukey HSD）
- 思想品德详细（~600 字）
- 模块用 `（1）（2）` 编号
- 下一步计划：日期范围 + 详细描述

### yl-模板（理论推导型）

**适用场景：**
- 学术学位
- 偏理论和算法的研究
- 公式密集（≥20 个）
- 核心贡献是"提出了一个算法/方法/模型"

**特点：**
- 直接进入详细展开（无总述和概述）
- 大量数学公式和推导
- 对比表格丰富
- 思想品德简洁（~200 字）
- 模块用 `1. 2.` 编号，内部用 `1.1, 1.2` 小节
- 下一步计划：编号列表 + 简短目标（无日期）

### 快速判断

| 判断维度 | lyw-模板 | yl-模板 |
|---------|---------|--------|
| 核心贡献 | 系统/平台/原型 | 算法/方法/模型 |
| PPT 特点 | 图片多、实验数据图表 | 公式多（≥20）、以推导为主 |
| 学位类型 | 专业学位 | 学术学位 |

## 文件结构

```
工作目录/
├── rewrite_report.py           # 生成脚本（主程序）
├── 模板.docx                   # 格式模板（封面页、表格结构）
├── 研究生成绩单.pdf             # 课程成绩、学分、绩点
├── 论文.pdf                    # 技术内容来源
├── 中期ppt.pdf                 # 工作框架来源
├── 中期考核相关资料.docx        # 考核要求
├── 图片/                        # 所有待插入报告的图片
│   ├── 整体框架.png
│   ├── 图01_xxx.png
│   ├── 图02_xxx.jpg
│   └── ...
├── ppt_images/                  # 从 PPT 提取的嵌入式图片（用于比对）
├── paper_images/                # 从论文提取的图片（用于比对）
├── 写入文件.docx                # 最终输出
├── 需求与约束.md                # 需求文档
├── modules_output.txt           # 自动检测的模块列表
├── SKILL_github.md              # Claude Code 技能文档
└── templates/                   # 模板参考文件
    ├── rewrite_report_template.py  # 脚本模板
    ├── lyw-模板.md                 # lyw 模板行文分析
    └── yl-模板.md                  # yl 模板行文分析
```

## 技术细节

### OMML 公式渲染

公式渲染流程：

```
LaTeX 字符串
    ↓
写入临时 .tex 文件
    ↓
pdflatex 编译（确保 LaTeX 语法正确）
    ↓
pandoc 转换为 .docx（LaTeX → OMML）
    ↓
提取 oMath/oMathPara 元素
    ↓
插入 Word 单元格（居中、无缩进）
```

**关键代码：**

```python
def add_formula(cell, formula_name: str):
    latex = FORMULAS[formula_name]
    # 1. 写入临时 .tex 文件
    with tempfile.NamedTemporaryFile(suffix='.tex', ...) as f:
        f.write(f'\\documentclass{{article}}\\begin{{document}}${latex}$\\end{{document}}')
    # 2. pdflatex 编译
    subprocess.run(['pdflatex', '-interaction=nonstopmode', tex_path])
    # 3. pandoc 转换
    subprocess.run(['pandoc', tex_path, '-o', docx_path, '--from=latex', '--to=docx'])
    # 4. 提取 oMath 元素
    temp_doc = Document(docx_path)
    for temp_p in temp_doc.paragraphs:
        for elem in temp_p._element:
            if elem.tag.endswith('}oMath') or elem.tag.endswith('}oMathPara'):
                p._element.append(elem)
```

### description 与公式分离规则

**description 中严禁出现数学表达式。** 所有公式必须通过 `add_formula` 渲染为 OMML 格式。

```python
# 错误：公式写入了 description
'核心公式为Delta y_k = Phi_k^T Delta C_k'

# 正确：描述含义，公式由 add_formula 渲染
'核心公式基于偏格式数据模型提取动态映射。'
# 然后调用 add_formula('mfac')
```

### 公式引入语规则

公式不能凭空出现，必须有文字引入：

```python
# 正确
'协同缺损向量定义为健康协同模式与患者残余协同之差的非负截断：'
# 然后插入 hdiff 公式

# 错误：公式凭空出现
'协同缺损评估方面，基于按需辅助机制...'
# 直接插入 hdiff 公式 — 突兀
```

### 图片内联插入规则

图片必须紧跟引用位置，不能集中堆放在模块末尾：

```python
Module(
    description='...如图2所示为VAF分析...\n不同阻力和疲劳因素...',
    images=['图04.jpeg', '图05.jpeg', '协同模式提取.png'],
    image_after={0: 0, 1: 1, 2: 3},  # 图片紧跟对应段落
)
```

## 常见问题

### Q1: 公式显示为纯文本

**原因：** description 中包含了数学表达式（如 `L_total = ...`）。

**解决：** 从 description 中移除所有数学表达式，仅用文字描述含义，公式由 `add_formula` 渲染。

### Q2: 图片重复使用

**原因：** 同一张图片被分配到多个模块。

**解决：** 检查所有 Module 的 `images` 列表，确保每张图片只出现一次。

### Q3: 图编号不连续

**原因：** `fig_num` 起始值错误或图片遗漏。

**解决：** 确保 `fig_num` 从 1 开始，且所有图片都被正确插入。

### Q4: pandoc 报错

**原因：** 未安装 pandoc 或 pdflatex。

**解决：**
```bash
# 检查 pandoc
where pandoc

# 检查 pdflatex
where pdflatex

# 安装
conda install -c conda-forge pandoc
# 或从官网下载 MiKTeX（含 pdflatex）
```

### Q5: 段落缩进不生效

**原因：** 缺少 `firstLineChars` 属性。

**解决：** 确保 `_apply_indent` 同时设置 `firstLineChars='200'`、`firstLine='480'`、`left='34'`。

### Q6: 正文两端对齐而非左对齐

**原因：** `add_para` 默认设了 `p.alignment = JUSTIFY`。

**解决：** 在 `add_para` 中显式设置 `p.alignment = WD_ALIGN_PARAGRAPH.LEFT`。

### Q7: 段落间距异常

**原因：** Word 默认段后间距不为 0。

**解决：** 在 `_apply_spacing` 中显式设置 `spacing.set(qn('w:after'), '0')`。

### Q8: formula_after/image_after 索引越界

**原因：** 段落数量变化后未更新索引。

**解决：** 运行 `validate_modules()` 自动检查索引范围。

## 开发说明

### 校验函数

```python
def validate_modules():
    """校验模块定义的正确性"""
    # 1. 检查 formula_after/image_after 索引范围
    # 2. 扫描 description 中的公式表达式残留
    # 3. 验证图片无重复使用
```

### 关键设计决策

1. **python-docx 操作表格单元格**：模板使用 6 行表格（奇数行=标题，偶数行=内容）
2. **pandoc 而非手写 OMML**：手写 OMML XML 容易出错，pandoc 转换更可靠
3. **pdflatex 预编译**：pandoc 直接处理 LaTeX 有时产生空 docx，pdflatex 预编译可解决
4. **图片目录作为唯一来源**：不从 PPT 或论文中提取图片，简化流程

## 许可证

MIT License

## 相关链接

- [python-docx 文档](https://python-docx.readthedocs.io/)
- [pandoc 官网](https://pandoc.org/)
- [MiKTeX 下载](https://miktex.org/download)
