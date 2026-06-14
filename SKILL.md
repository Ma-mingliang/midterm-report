---
name: midterm-report
description: 研究生中期考核报告自动生成技能。基于论文PDF、中期PPT和成绩单，自动分析研究内容，动态生成符合学校格式要求的中期考核报告Word文档。支持任意研究方向，自适应模块检测、三层行文逻辑、中英混合字体、OMML公式，输出≥6000字、≥15张图片的完整报告。
---

# 中期考核报告自动生成技能（通用版）

基于 python-docx 自动生成研究生中期考核报告（写入文件.docx），涵盖思想品德自述、已完成科研工作、下一步科研计划三个章节。**适用于任意研究方向**，通过动态分析 PPT 和论文自动识别研究模块。

---

## 必需输入文件

| 文件 | 用途 | 能否省略 |
|------|------|---------|
| `研究生成绩单.pdf` | 课程成绩、学分、绩点 | **不可省略** |
| 论文PDF | 技术内容来源 | **不可省略** |
| `中期ppt.pdf` | 工作框架来源 | **不可省略** |
| `图片/` 目录 | 用户提供的图片（可选） | 可省略（从论文/PPT提取） |

**不需要的文件：**
- ~~模板.docx~~ — 格式规范已内置
- ~~中期考核相关资料.docx~~ — 考核要求已内置
- ~~参考文档~~ — 图片验证通过视觉模型完成

---

## 核心设计：动态模块检测

本技能的核心能力是**从 PPT 和论文中自动识别研究模块结构**，而非依赖预设模板。

### 模块检测流程

```
阶段1: PPT页面分析
  ├─ 提取每页标题文字
  ├─ 识别页面类型（封面/目录/内容/总结/致谢）
  ├─ 检测层级结构（大标题 → 子标题 → 内容）
  └─ 输出：页面分类表

阶段2: 模块边界识别
  ├─ 大标题页面 = 模块分界点
  ├─ 目录页面 = 模块清单
  ├─ 连续内容页 = 同一模块
  └─ 输出：模块列表（标题 + 页码范围 + 图片）

阶段3: 内容提取
  ├─ 每个模块提取：方法描述、公式、图片、实验数据
  ├─ 从论文中补充：技术细节、引用、对比实验
  └─ 输出：模块内容摘要

阶段4: 图片映射
  ├─ PPT嵌入式图片 → 按模块分组
  ├─ 论文图表 → 按章节分组
  ├─ 交叉验证：视觉模型 + PPT文字 + 论文描述
  └─ 输出：图片映射表（文件名 → 模块 → 图注）
```

### PPT 页面分类规则

| 页面特征 | 分类 | 处理方式 |
|----------|------|---------|
| 首页/含"中期"/"汇报" | 封面 | 提取标题、作者、日期 |
| 含"目录"/"Contents"/编号列表 | 目录 | 提取模块清单 |
| 含"参考文献"/"References" | 参考文献 | 跳过 |
| 含"致谢"/"Thanks" | 致谢 | 跳过 |
| 大标题 + 内容 | 内容页 | 归入当前模块 |
| 仅大标题（无子内容） | 模块分界 | 开启新模块 |

### 模块数量自适应

- **最少模块数**：3（基础研究：方法 + 实验 + 结论）
- **典型模块数**：4-6（常规研究）
- **最多模块数**：8（复杂研究）
- 如果检测到超过8个模块，合并相近的小模块

---

## 模板选择

用户可选择两种行文与排版模板，生成报告前必须确认使用哪一种。模板内容已内置在 `templates/` 目录中。

### lyw-模板（工程实现型）
- **特点**：三层行文逻辑（总述→概述→详述）、课程成绩逐门列出、偏工程实现和实验验证、统计分析丰富
- **适用**：工程类专业学位、偏实验验证的研究、图片数量多（≥15张）
- **思想品德**：详细（~600字，列出每门课程和分数）
- **科研工作**：三层逻辑展开，模块用（1）（2）编号
- **下一步计划**：日期范围+详细描述

### yl-模板（理论推导型）
- **特点**：直接进入详细展开、大量数学公式和推导、每个模块独立实验验证、对比表格丰富
- **适用**：学术学位、偏理论和算法的研究、公式密集（≥20个）
- **思想品德**：简洁（~200字，不列具体课程）
- **科研工作**：直接展开，模块用1. 2. 编号，内部用1.1, 1.2小节
- **下一步计划**：编号列表+简短目标（无日期）

### 选择依据

向用户说明两种模板的特点，询问其研究更偏向理论创新还是工程实现。

**快速判断：**
- 核心贡献是"做了一个系统/平台/原型" → lyw-模板
- 核心贡献是"提出了一个算法/方法/模型" → yl-模板

**材料分析辅助：**
- PPT中图片多（≥15张）且包含大量实验数据图表 → lyw-模板
- PPT中公式多（≥20个）且以推导为主 → yl-模板

---

## 工作流程（6阶段）

### 阶段1：需求收集与材料分析

**必做步骤：**

1. **扫描材料清单**：
   - `研究生成绩单.pdf` — 课程成绩、学分、绩点（**必需**）
   - 论文PDF — 技术内容来源（**必需**）
   - `中期ppt.pdf` — 完整工作框架（**必需**）
   - `图片/` 目录 — 用户提供的图片（可选）

2. **读取成绩单**：提取总学分、绩点、核心课程及分数

3. **分析PPT结构**：
   - 提取每页文字内容（前200字）
   - 识别页面类型（封面/目录/内容/总结）
   - 检测模块边界（大标题页面）
   - 输出模块清单

4. **确认信息来源**：向用户确认检测到的模块结构是否正确

### 中期考核要求（内置）

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
| 二、已完成的科研工作 | 整体框架 + N个技术模块详细描述 | ~60% |
| 三、下一步科研计划 | 分阶段科研安排 | ~20% |

#### 第二章三层行文逻辑（lyw-模板，必须遵循）

**第一层：开篇总述**
- 研究目的（1-2句，从PPT首页/摘要提取）
- 整体框架描述（配框架图，从PPT中提取）
- "目前的已完成工作主要包括以下N个方面："

**第二层：N部分简要概述**
- 每个模块用1-2句话概括核心内容
- 不展开细节，仅提供全局视角
- 编号：（1）（2）（3）...
- 示例（通用格式）：
  ```
  （N）[模块名称]，[做了什么]，[怎么做的]，[达到什么效果/为后续提供什么]。
  ```

**第三层：详细展开**
- "已完成工作：" 过渡标记
- 每个模块独立展开，包含：
  - 方法描述（技术细节、公式）
  - 图片展示（1-3张/模块）
  - 实验验证（数据、统计分析）
  - 创新点说明

#### 第二章行文逻辑（yl-模板，直接展开）

- 无开篇总述和简要概述
- 直接"已完成工作："后进入各模块
- 模块用 1. 2. 编号
- 内部用 1.1, 1.2 小节
- 大量数学公式和推导

### 阶段4：脚本生成

**技术栈：** python-docx + PyMuPDF + Pillow + omml_formulas

**核心函数模式：**（详见模板文件 `templates/rewrite_report_template.py`）

**图片路径：**
```python
BASE_DIR = r'工作目录路径'
IMG_DIR = os.path.join(BASE_DIR, '图片')  # 用户提供的图片（可选）

def img(name):
    """图片路径查找（优先级：图片/ → ppt_images/ → paper_images/）"""
    # 详见 rewrite_report_template.py
```

**模板格式规范：**

| 元素 | 字体 | 字号 | 缩进 | 对齐 |
|------|------|------|------|------|
| 章节标题 | 仿宋_GB2312 | 15pt | 无 | 左对齐 |
| 正文段落 | 宋体+Times New Roman | 12pt | firstLine=480 | 两端对齐 |
| 图注 | Times New Roman | 10pt | firstLine=400 | 居中 |

所有段落行距：`line=360, lineRule=auto, after=0`

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
| 标题 | 仿宋_GB2312 | — | 15pt |
| 图注 | 宋体 | Times New Roman | 10pt（五号） |

### 编号
- 小节编号：（1）（2）（3）...
- 图片编号：图1 图2 图3...（连续无跳号）

### 图片
- 居中放置
- 图注在图片下方
- 宽度建议 4.5-5.5 英寸
- 图注段落样式：a3

---

## 内容风格指南

### 语言要求
- 使用"本人"等第一人称
- 正式学术报告风格
- 逻辑清晰，层次分明

### 技术深度
- 包含关键公式（损失函数、控制律、目标函数等）
- 描述核心方法思路和创新点
- 不展开详细推导过程
- 适度引用实验数据和统计结果

### 篇幅控制
- 思想品德章节：6-8段（lyw）或 1段（yl）
- 科研工作章节：每个模块2-4段
- 下一步计划：分3个时间段，每段1段描述

---

## 通用写作模式

### 模块展开模式（每个技术模块通用）

```
1. 背景与目的（为什么要做）
   "针对[问题]，本研究提出了[方法名称]..."

2. 方法描述（怎么做）
   "具体而言，[方法步骤1]，[方法步骤2]..."
   [配公式]

3. 图片展示（配图说明）
   "如图X所示为[内容描述]。"
   [配图片]

4. 实验验证（数据支撑）
   "实验结果表明，[结论]。如图X所示..."
   [配数据图/表格]

5. 创新点/结论
   "该方法的创新点在于[创新点描述]。"
```

### 问题导向模式
```
针对[问题]，设计[方法]，实现[目标]
```

### 流程式描述
```
首先[步骤1]，然后[步骤2]，最后[步骤3]
```

### 目的驱动模式
```
为了[目的]，[做了什么]
```

### 数据支撑模式
```
[方法]在[条件]下达到[数值]，较[基线]提升[百分比]
```

### 公式引入模式
```
[描述]定义为：
[公式]
其中[变量说明]。
```

### 图片引入模式
```
如图X所示为[内容描述]
```

---

## 常见问题与解决方案

### Q1: Python执行报错 "exit code 49"
**解决：** 使用完整路径 `E:/Anaconda/python.exe <file>`

### Q2: 图片图注与实际内容不匹配
**解决：** 必须通过阶段2的交叉验证流程确认图片真实含义，不能仅凭文件名判断

### Q3: 字数不足6000
**解决：** 补充技术细节：数据预处理流程、方法论描述、实验统计分析、稳定性分析、实验平台技术参数

### Q4: 中英混合字体设置不生效
**解决：** 使用 set_run_font 函数，确保 eastAsia、ascii、hAnsi 三个属性都设置

### Q5: 表格单元格内容残留
**解决：** 使用 clear_cell() 先清空再写入

### Q6: conda run 不支持多行Python代码
**解决：** 始终将Python代码写入文件后执行

### Q7: 段落行间距与模板不一致
**解决：** 在 add_para 和 add_caption 中显式设置 `line=360, lineRule=auto, after=0`

### Q8: PPT模块检测不准确
**解决：** 手动向用户确认模块结构，根据反馈调整模块划分

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

### 行文逻辑三层结构

模板的写作逻辑是"总述→概述→详述"三层结构：
1. **总述**：研究目的 + 整体框架 + 框架图
2. **概述**：N个模块各1-2句话概括
3. **详述**："已完成工作：" + 每个模块详细展开

缺少任何一层都会导致行文逻辑不完整。

### Windows环境注意事项

- `python` 命令可能返回 exit code 49，使用完整路径 `E:/Anaconda/python.exe`
- `conda run -n base python -c "多行代码"` 会报错，必须写入文件
- 路径中包含中文字符时，使用原始字符串 `r'路径'`

---

**技能版本：** 4.0.0
**兼容平台：** Claude Code (CLI/Desktop/Web)
**Python版本：** >= 3.7
**依赖：** python-docx, PyMuPDF (fitz), Pillow, lxml
**Python环境：** Anaconda（推荐）

<!-- SLOW_UPDATE_START -->
When generating the midterm report, always follow these priorities:

1. **Dynamic Module Detection — Core Capability**: The skill must detect research modules from PPT structure, not use hardcoded modules. Extract page titles, identify module boundaries (new major topic = new module), and build the chapter structure dynamically. If detection fails, ask the user to confirm the module list before proceeding.

2. **Chapter 3 Research Plan — Prevent Regression**: The '下一步科研计划' chapter requires as much care as Chapter 2. Structure it as 3 distinct time phases. Each phase must: (a) specify a clear date range, (b) describe 2-3 concrete research tasks with technical detail, (c) explicitly connect to the completed work from Chapter 2. Aim for ~20% of total report length (at least 1200 words).

3. **Image Cross-Validation is Non-Negotiable**: Never skip cross-validation for any image. Before inserting, verify meaning via visual inspection + PPT text + paper descriptions. Use the mapping table to track source, content, caption assignment, and verification status. If the visual model cannot clearly identify the image content, mark it as '待确认' and seek user input before inserting — do not guess.

4. **Three-Layer Logic for Chapter 2 (lyw-模板)**: Always include: (1) opening overview with framework figure, (2) brief summary of all N modules (1-2 sentences each), (3) detailed expansion after '已完成工作:' transition marker. Missing any layer causes structural failures.

5. **Template Selection**: Analyze formula count and image count in provided materials to recommend lyw- (≥15 images, experimental) vs yl- (≥20 formulas, theoretical). Default to lyw- if uncertain.

6. **Caption and Format Consistency**: After generating, visually verify every image caption in the final document. Check for duplicate images, mismatched descriptions, and non-sequential figure numbering. Ensure all paragraphs use line=360, lineRule=auto, after=0.

7. **Execution**: Use E:/Anaconda/python.exe with full path. Write all Python code to files before execution. Verify output has ≥6000 words and ≥15 images.
<!-- SLOW_UPDATE_END -->
