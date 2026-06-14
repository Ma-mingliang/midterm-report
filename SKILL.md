---
name: midterm-report
description: 研究生中期考核报告自动生成技能。基于论文PDF、中期PPT和成绩单，自动生成符合学校格式要求的中期考核报告Word文档。支持中英混合字体、OMML公式、图片来源追溯与交叉验证、三层行文逻辑结构，输出≥6000字、15-20张研究图片的完整报告。
---

# 中期考核报告自动生成技能

基于 python-docx 自动生成研究生中期考核报告（写入文件.docx），涵盖思想品德自述、已完成科研工作、下一步科研计划三个章节。

---

## 必需输入文件

| 文件 | 用途 | 能否省略 |
|------|------|---------|
| `研究生成绩单.pdf` | 课程成绩、学分、绩点 | **不可省略** |
| `论文.pdf` | 技术内容来源 | **不可省略** |
| `中期ppt.pdf` | 工作框架来源 | **不可省略** |
| `图片/` 目录 | 用户提供的图片（可选） | 可省略（从论文/PPT提取） |

**不需要的文件：**
- ~~模板.docx~~ — 格式规范已内置
- ~~中期考核相关资料.docx~~ — 考核要求已内置
- ~~写入文件(3).docx~~ — 图片验证不再依赖参考文档

---

## 实际工作状态

用户的工作模式：

1. **论文和PPT是核心参考**：用户会提供论文PDF和中期PPT，这些是写作内容的主要来源。需要从中提取技术细节、实验数据、方法描述等。
2. **图片来源**：优先从PPT中提取**嵌入式图片**，其次从论文中提取图表。用户也可直接提供 `图片/` 目录。PPT整页渲染仅用于理解页面结构，嵌入式图片才是实际使用的图表。
3. **必须理解图片的真实含义**：不能仅凭文件名或位置猜测图片用途。必须通过阅读PPT页面文字、论文上下文来确认每张图片**实际表达的内容**，并与图注严格对应。必须交叉验证，确保正确无误。

---

## 模板选择

用户可选择两种行文与排版模板，生成报告前必须确认使用哪一种。模板内容已内置在 `templates/` 目录中，无需外部模板文件。

### lyw-模板（工程实现型）
- **文件**：`templates/lyw-模板.md`（内置）
- **特点**：三层行文逻辑（总述→概述→详述）、课程成绩逐门列出、偏工程实现和实验验证、统计分析丰富（ANOVA、Tukey HSD）
- **适用**：工程类专业学位、偏实验验证的研究、图片数量多（15-20张）
- **思想品德**：详细（~600字，列出每门课程和分数）
- **科研工作**：三层逻辑展开，模块用（1）（2）编号
- **下一步计划**：日期范围+详细描述

### yl-模板（理论推导型）
- **文件**：`templates/yl-模板.md`（内置）
- **特点**：直接进入详细展开、大量数学公式和推导、每个模块独立实验验证、对比表格丰富
- **适用**：学术学位、偏理论和算法的研究、公式密集（25+个）
- **思想品德**：简洁（~200字，不列具体课程）
- **科研工作**：直接展开，模块用1. 2. 编号，内部用1.1, 1.2小节
- **下一步计划**：编号列表+简短目标（无日期）

### 选择方式
在阶段1中询问用户选择哪种模板，然后严格按照对应模板的行文逻辑组织内容。

**选择依据建议：**
1. **研究类型判断**：
   - **yl-模板**：适用于理论推导、算法创新、公式密集（≥20个）、数学模型构建为主的研究。如控制理论、算法设计、理论分析等。
   - **lyw-模板**：适用于工程实现、系统开发、实验验证为主、图片/数据图表丰富（≥15张）的研究。如系统设计、原型开发、实验对比等。
2. **用户确认**：向用户说明两种模板的特点，询问其研究更偏向理论创新还是工程实现，并确认模板选择。
3. **默认策略**：若用户无法明确判断，可先询问研究的核心贡献是理论突破还是系统实现，再推荐相应模板。
4. **材料分析辅助**：自动分析用户提供的材料（如PPT、论文）中的公式数量、图片数量、实验类型等，为模板选择提供数据支持。例如，若材料中公式≥20个且图片≤15张，可推荐yl-模板；若图片≥15张且包含大量实验数据图表，可推荐lyw-模板。

---

## 工作流程（6阶段）

### 阶段1：需求收集与材料分析

**必做步骤：**

1. **扫描材料清单**：
   - `研究生成绩单.pdf` — 课程成绩、学分、绩点（**必需**）
   - 论文PDF — 技术内容来源（**必需**）
   - `中期ppt.pdf` — 完整工作框架（**必需**）
   - `图片/` 目录 — 用户提供的图片（可选，如无则从论文/PPT提取）
2. **确认信息来源**：向用户确认各材料的用途和对应关系

### 中期考核要求（内置）

以下为研究生中期考核的通用要求，无需外部文件：

**考核内容：**
1. 思想品德与业务学习情况自述
2. 已完成的科研工作（主要内容）
3. 下一步科研计划

**评分标准：**
- 思想品德：政治思想、学习态度、学术道德、身体素质
- 科研工作：研究内容完整性、技术深度、创新性、实验验证、成果产出
- 科研计划：目标明确性、可行性、与已完成工作的衔接

**格式要求：**
- 总字数 ≥ 6000字
- 图片数量 ≥ 15张
- 图片编号连续无跳号
- 中英混合字体（宋体+Times New Roman）
- 1.5倍行距，段后间距0

### 阶段2：图片提取与理解（关键阶段）

**目标：获取所有需要插入报告的图片，并通过交叉验证确认每张图片的真实含义和用途。**

**图片来源优先级：**
1. **PPT嵌入式图片**（最可靠）— 从 `中期ppt.pdf` 提取
2. **论文图片**（较可靠）— 从 `论文.pdf` 提取
3. **用户提供的图片**（需验证）— 从 `图片/` 目录获取（如有）

**这是最容易出错的阶段。必须严格遵循以下流程。**

#### 2.1 从PPT中提取嵌入式图片

**重要：PPT中需要提取的是嵌入式图片，不仅仅是整页渲染。**

```python
import fitz
import os

doc = fitz.open('中期ppt.pdf')

# 提取嵌入式图片（实际使用的图表、数据图等）
for i in range(doc.page_count):
    page = doc[i]
    images = page.get_images(full=True)
    for j, img in enumerate(images):
        xref = img[0]
        base_image = doc.extract_image(xref)
        ext = base_image['ext']
        data = base_image['image']
        with open(f'ppt_images/page{i+1}_img{j}.{ext}', 'wb') as f:
            f.write(data)
        print(f'Page {i+1}, img{j}: {base_image["width"]}x{base_image["height"]}, {len(data)} bytes')

doc.close()
```

**整页渲染仅用于理解页面结构和布局，不直接用于报告：**

```python
doc = fitz.open('中期ppt.pdf')
for i in range(doc.page_count):
    page = doc[i]
    pix = page.get_pixmap(dpi=150)
    pix.save(f'ppt_images/page{i+1}_render.png')
doc.close()
```

#### 2.2 从论文PDF中提取图片

```python
doc = fitz.open('论文.pdf')
for i in range(doc.page_count):
    page = doc[i]
    images = page.get_images(full=True)
    for j, img in enumerate(images):
        xref = img[0]
        base_image = doc.extract_image(xref)
        with open(f'paper_images/paper_p{i+1}_img{j}.{base_image["ext"]}', 'wb') as f:
            f.write(base_image['image'])
        print(f'Paper page {i+1}, img{j}: {base_image["width"]}x{base_image["height"]}')
doc.close()
```

#### 2.3 提取PPT每页文字内容（用于理解图片含义）

```python
import fitz

doc = fitz.open('中期ppt.pdf')
for i in range(doc.page_count):
    page = doc[i]
    text = page.get_text()[:200].replace('\n', ' ').strip()
    print(f'Page {i+1}: {text}')
doc.close()
```

#### 2.4 理解图片真实含义并交叉验证（必须执行）

**核心原则：必须通过交叉验证确认每张图片的真实含义。不能仅凭文件名、位置或猜测。**

**必须执行的验证步骤：**

1. **用视觉模型查看每张图片**：描述图片的实际内容（图表类型、坐标轴、数据趋势、标注文字等）
2. **阅读PPT对应页面的完整文字**：理解该页面的主题、上下文、图片在页面中的作用
3. **阅读论文中对应的章节**：找到论文中对该图表的描述和解释
4. **交叉比对**：
   - 图片中的标注文字是否与PPT/论文描述一致？
   - 图片展示的数据/方法是否与上下文匹配？
   - 图注描述是否准确反映了图片内容？
5. **建立映射表**：记录每张图片的：
   - 文件名
   - 来源（PPT第X页第X张嵌入图 / 论文第X页第X张 / 用户直接提供）
   - 实际内容描述（从视觉模型+PPT文字+论文描述综合得出）
   - 对应的报告图注（图1、图2...）
   - 验证状态（已验证/待确认）

**注意：不再依赖参考文档进行图片验证。所有验证通过视觉模型+PPT文字+论文描述完成。**

**交叉验证示例：**

```
图片文件：ppt_images/page6_img0.png
视觉模型描述：折线图，横轴为epoch，纵轴为VAF，显示3条曲线趋于收敛
PPT第6页文字：协同重构精度验证，VAF在r=3条件下达到90%以上
论文第4.2节：Figure 5 shows the VAF convergence...
验证结论：此图为VAF收敛曲线，对应报告图4"VAF与协同数的关系" ✓
```

**常见错误（必须避免）：**
- 仅凭文件名猜测图片内容（文件名可能完全错误）
- 同一图片被错误地用于多个图注
- 图注描述与图片实际内容不符
- PPT整页渲染图与嵌入式图片混淆

### 阶段3：内容规划与行文逻辑

**三个章节结构：**

| 章节 | 内容 | 篇幅占比 |
|------|------|---------|
| 一、思想品德与业务学习情况自述 | 思想政治、课程成绩、学习态度、身体素质 | ~20% |
| 二、已完成的科研工作 | PPT整体框架 + 6个技术模块详细描述 | ~60% |
| 三、下一步科研计划 | 分阶段科研安排 | ~20% |

#### 第二章三层行文逻辑（必须遵循）

**第一层：开篇总述**
- 研究目的（1-2句）
- 整体框架描述（配框架图）
- "目前的已完成工作主要包括以下六个方面："

**第二层：六部分简要概述**
- 每个模块用1-2句话概括核心内容
- 不展开细节，仅提供全局视角
- 示例：
  ```
  （1）多模态数据采集与预处理，搭建多模态数据采集平台，采集sEMG和关节运动学数据，
  并进行信号滤波、包络提取和标准化等预处理操作，为后续协同建模提供高质量数据基础。
  ```

**第三层：详细展开**
- "已完成工作：" 过渡标记
- 每个模块独立展开，包含：
  - 方法描述（技术细节、公式）
  - 图片展示（1-3张/模块）
  - 实验验证（数据、统计分析）
  - 创新点说明

**6个技术模块（按PPT顺序）：**

| 模块 | 内容 | 典型图片数 |
|------|------|-----------|
| （1）多模态数据采集与预处理 | 采集平台、肌肉定义、预处理流程、VAF分析 | 2-3张 |
| （2）跨模态泛化协同建模 | NMF特征提取、ResMamba网络、知识蒸馏、迁移学习、实验验证 | 6-8张 |
| （3）实时协同缺损评估与前馈解码 | AAN机制、协同缺损定义、滑动窗口评估 | 1-2张 |
| （4）基于MFAC的融合闭环控制律 | PFDM数据驱动、前馈补偿、融合控制律、稳定性分析 | 1-2张 |
| （5）实验平台搭建 | 三个子系统、时序同步、技术参数 | 1-2张 |
| （6）学术成果 | 论文、软著、专利、比赛 | 1张 |

### 阶段4：脚本生成

**技术栈：** python-docx + PyMuPDF + Pillow + omml_formulas

**核心函数模式：**（详见模板文件 `templates/rewrite_report_template.py`）

**图片路径：**
```python
BASE_DIR = r'工作目录路径'
IMG_DIR = os.path.join(BASE_DIR, '图片')  # 用户提供的图片（可选）

def img(name):
    """图片路径（优先从图片/目录，其次从PPT/论文提取目录）"""
    # 首先检查用户提供的图片目录
    user_img_path = os.path.join(IMG_DIR, name)
    if os.path.exists(user_img_path):
        return user_img_path

    # 其次检查PPT提取的图片目录
    ppt_img_path = os.path.join(BASE_DIR, 'ppt_images', name)
    if os.path.exists(ppt_img_path):
        return ppt_img_path

    # 最后检查论文提取的图片目录
    paper_img_path = os.path.join(BASE_DIR, 'paper_images', name)
    if os.path.exists(paper_img_path):
        return paper_img_path

    # 如果都找不到，返回用户目录路径（会让调用者知道图片缺失）
    return user_img_path
```

**模板格式规范：**

| 元素 | 字体 | 字号 | 缩进 | 对齐 |
|------|------|------|------|------|
| 章节标题 | 仿宋_GB2312 | 15pt | 无 | 左对齐 |
| 正文段落 | 宋体+Times New Roman | 12pt | firstLine=480 | 两端对齐 |
| 图注 | Times New Roman | 10pt | firstLine=400 | 居中 |

所有段落行距：`line=360, lineRule=auto, after=0`

**OMML公式（pandoc LaTeX → OMML）：**
```python
from omml_formulas import create_formula_paragraph
create_formula_paragraph(cell, 'loss')
create_formula_paragraph(cell, 'pfdm')
```

### 阶段5：执行与验证

```bash
E:/Anaconda/python.exe "D:/工作目录/rewrite_report.py"
```

**验证检查清单：**
- [ ] 文字量 >= 6000字
- [ ] 图片数量 >= 15张
- [ ] 图片编号连续无跳号
- [ ] 每张图片内容与图注一致（已通过阶段2的交叉验证）
- [ ] 公式居中显示
- [ ] 字体格式与模板一致
- [ ] 内容与中期考核要求对齐（思想品德、科研工作、科研计划三章）
- [ ] 成绩数据与研究生成绩单一致

### 阶段6：图注二次核验

**用视觉模型逐张核验图片内容与图注是否匹配。**

1. 读取文档中所有插入的图片
2. 对每张图片描述其实际内容
3. 与图注文字对比
4. 修正不匹配的图注

**常见图注错误：**
- 图片重复使用（同一张图出现在两个位置）
- 图注描述与图片实际内容不符
- 图片顺序错乱导致编号混乱

---

## 格式规范

### 字体
| 元素 | 中文字体 | 英文字体 | 字号 |
|------|---------|---------|------|
| 正文 | 宋体 | Times New Roman | 12pt（小四） |
| 标题 | 黑体 | — | — |
| 图注 | 宋体 | Times New Roman | 10pt（五号） |

### 编号
- 小节编号：（1）（2）（3）...
- 图片编号：图1 图2 图3...（连续无跳号）

### 图片
- 居中放置
- 图注在图片下方
- 宽度建议 4.5-5.5 英寸
- 图注段落样式：a3（与模板一致）

---

## 内容风格指南

### 语言要求
- 使用"本人"等第一人称
- 正式学术报告风格
- 逻辑清晰，层次分明

### 技术深度
- 包含关键公式（损失函数、控制律等）
- 描述核心方法思路和创新点
- 不展开详细推导过程
- 适度引用实验数据和统计结果

### 篇幅控制
- 思想品德章节：6-8段
- 科研工作章节：每个模块2-4段
- 下一步计划：分3个时间段，每段1段描述

---

## 常见问题与解决方案

### Q1: Python执行报错 "exit code 49"
**解决：** 使用完整路径 `E:/Anaconda/python.exe <file>`

### Q2: 图片图注与实际内容不匹配
**解决：** 必须通过阶段2的交叉验证流程确认图片真实含义，不能仅凭文件名判断

### Q3: 字数不足6000
**解决：** 补充技术细节：数据预处理流程、NMF/VAF方法论、实验统计分析、控制律稳定性分析、实验平台技术参数

### Q4: 中英混合字体设置不生效
**解决：** 使用 set_run_font 函数，确保 eastAsia、ascii、hAnsi 三个属性都设置

### Q5: 表格单元格内容残留
**解决：** 使用 clear_cell() 先清空再写入

### Q6: conda run 不支持多行Python代码
**解决：** 始终将Python代码写入文件后执行

### Q7: 段落行间距与模板不一致
**解决：** 在 add_para 和 add_caption 中显式设置 `line=360, lineRule=auto, after=0`

---

## 文件结构

```
工作目录/
├── 研究生成绩单.pdf        # 成绩数据（必需）
├── 论文.pdf               # 技术内容来源（必需）
├── 中期ppt.pdf            # 工作框架来源（必需）
├── 图片/                   # 用户提供的图片（可选，如无则从论文/PPT提取）
├── ppt_images/             # 从PPT提取的嵌入式图片（自动生成）
├── paper_images/           # 从论文提取的图片（自动生成）
├── omml_formulas.py        # OMML公式构建模块（自动生成）
├── rewrite_report.py       # 生成脚本（自动生成）
└── 写入文件.docx            # 最终输出（自动生成）
```

---

## 关键经验教训

### 图片理解优先级

1. **最可靠**：视觉模型查看图片实际内容 + PPT/论文文字描述的交叉验证
2. **较可靠**：从PPT中提取的嵌入式图片（需验证页码和内容）
3. **较可靠**：从论文中提取的图片（需验证上下文）
4. **不可靠**：仅凭文件名判断图片内容
5. **不可靠**：PPT整页渲染图（可能包含多个元素，不能直接用于报告）

**注意：不再依赖参考文档进行图片验证。所有验证通过视觉模型+PPT文字+论文描述完成。**

### 行文逻辑三层结构

模板的写作逻辑是"总述→概述→详述"三层结构：
1. **总述**：研究目的 + 整体框架 + 框架图
2. **概述**：6个模块各1-2句话概括
3. **详述**："已完成工作：" + 每个模块详细展开

缺少任何一层都会导致行文逻辑不完整。

### Windows环境注意事项

- `python` 命令可能返回 exit code 49，使用完整路径 `E:/Anaconda/python.exe`
- `conda run -n base python -c "多行代码"` 会报错，必须写入文件
- 路径中包含中文字符时，使用原始字符串 `r'路径'`

---

**技能版本：** 3.1.0
**兼容平台：** Claude Code (CLI/Desktop/Web)
**Python版本：** >= 3.7
**依赖：** python-docx, PyMuPDF (fitz), Pillow, lxml
**Python环境：** Anaconda（推荐）

<!-- SLOW_UPDATE_START -->
When generating the midterm report, always follow these priorities:

1. **Chapter 3 Research Plan — Prevent Regression**: The '下一步科研计划' chapter requires as much care as Chapter 2. Structure it as 3 distinct time phases (e.g., 2024.09-2025.03, 2025.03-2025.09, 2025.09-2026.06). Each phase must: (a) specify a clear date range, (b) describe 2-3 concrete research tasks with technical detail, (c) explicitly connect to the completed work from Chapter 2. For lyw-模板, use detailed descriptions with date ranges per phase; for yl-模板, use numbered sub-items with concise goals. The research plan must logically follow from and build upon what was already completed — do not introduce unrelated directions. Ensure the plan spans the remaining degree period and includes thesis writing milestones. Aim for ~20% of total report length (at least 1200 words).

2. **Image Cross-Validation is Non-Negotiable**: Never skip cross-validation for any image. Before inserting, verify meaning via visual inspection + PPT text + paper descriptions. Use the mapping table to track source, content, caption assignment, and verification status. If the visual model cannot clearly identify the image content (e.g., too small, ambiguous, or no readable text), mark it as '待确认' and seek user input before inserting — do not guess. Even if the report passes word-count checks, mismatched images will fail quality evaluation.

3. **Three-Layer Logic for Chapter 2**: Always include: (1) opening overview with framework figure, (2) brief summary of all 6 modules (1-2 sentences each), (3) detailed expansion after '已完成工作:' transition marker. Missing any layer causes structural failures.

4. **Template Selection**: Analyze formula count and image count in provided materials to recommend lyw- (≥15 images, experimental) vs yl- (≥20 formulas, theoretical). Default to lyw- if uncertain.

5. **Caption and Format Consistency**: After generating, visually verify every image caption in the final document. Check for duplicate images, mismatched descriptions, and non-sequential figure numbering. Ensure all paragraphs use line=360, lineRule=auto, after=0. For yl-模板, verify all OMML formulas render correctly and are centered in their cells.

6. **Execution**: Use E:/Anaconda/python.exe with full path. Write all Python code to files before execution. Verify output has ≥6000 words and ≥15 images.
<!-- SLOW_UPDATE_END -->
