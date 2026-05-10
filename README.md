# paper2ppt

中文 | [English](#english)

## 中文

`paper2ppt` 收录了一个面向论文汇报 PPT 生成的 Codex Skill: `uav-paper-report`。

它不是一个固定模板的“一键论文转 PPT”产品，而是一套可复用的生成与审查经验：让 Codex 在阅读论文、理解 PPT 模板、组织汇报逻辑、排版公式图表、检查视觉问题时，尽量保持稳定的论文组会汇报质量。

这个 skill 目前主要服务于无人机、机器人、轨迹规划、SLAM、控制、自主系统等方向的中文论文汇报。它也可以作为其它学术方向的 PPT 生成工作流参考。

### 这个 Skill 解决什么问题

在真实的论文汇报里，最难的通常不是“把论文摘要复制到 PPT 里”，而是这些细节：

- 给定一篇论文后，需要提炼出可以讲清楚的背景、问题、方法、公式、实验和局限。
- 给定一个已有 PPT 模板后，需要沿用它的标题层级、颜色、页眉、线条、表格、公式区域和结尾页风格。
- 方法部分不能太空，也不能把一页塞爆；公式要能配合讲解，而不是像截图一样贴上去。
- 图不能随便拉伸、裁断或和文字重叠；图旁边需要说明它在证明什么。
- 正文字号、缩进、项目符号、红色强调、段落间距要统一，不能一页一个样。
- 生成后需要真的渲染出来检查，而不是只看 PPTX 文件结构。

`uav-paper-report` 把这些要求写成了可执行的流程、脚本和样例，方便 Codex 在类似任务里重复使用。

### 适合的使用场景

这个 skill 适合下面这些任务：

- 读一篇无人机或机器人相关论文，生成中文组会汇报 PPT。
- 按照用户给的 PPT 模板或历史汇报样式生成新论文汇报。
- 把现有生成结果继续改到更像人工整理的论文汇报。
- 检查 PPT 是否存在文字重叠、图片比例错误、空白过多、无意义换行、字号不统一等问题。
- 对多个模板和多篇论文做泛化测试，看看生成流程是否稳定。
- 发布公开示例前清理 PPTX 中的个人信息、备注、批注和文档元数据。

如果你只是想找一个网页应用上传 PDF 后立刻下载 PPT，这个仓库还不是那种产品形态。它更像是给 Codex 使用的“论文汇报制作规范 + 可复用脚本包”。

### 能生成什么样的内容

默认目标是一份 18-24 页左右的中文论文汇报，结构通常包括：

- 封面和论文基本信息
- 研究背景与任务动机
- 相关工作或已有方法的不足
- 论文核心思路和技术路线
- 方法细节、公式解释和算法流程
- 关键实验设置、指标和结果表格
- 消融实验、对比实验或可视化结果
- 方法局限、适用条件和总结
- 统一风格的结束页

内容密度可以按需求调整。比如可以要求：

- “方法部分详细一点，背景简单一点”
- “公式多写一些，但实验结果压缩”
- “控制在 15 页以内”
- “做成 25 页左右，适合 30 分钟汇报”
- “多介绍实验设置，少写相关工作”

skill 的重点不是机械扩写，而是把每页做成适合讲给别人听的内容：每一页都有明确的结论、证据和解释。

### 模板适配方式

当用户提供 PPT 模板或历史汇报时，推荐的做法是先观察模板，再生成内容：

1. 读取 PPT 尺寸、封面、结束页、页眉、横线和主色。
2. 识别标题、一级正文、二级正文、表格、注释、公式区域的字号层级。
3. 判断模板更适合纵向堆叠、左右对照、图文并排还是表格解释。
4. 根据论文内容选择页面类型，而不是把所有页都套进同一种版式。
5. 生成后导出 PDF 和逐页 PNG，再检查是否真的没有重叠、溢出和空白过多。

公开仓库里不包含私人模板，也不包含任何个人汇报材料。这里保留的是 AI 生成的示例、截图、脚本和可公开的配置。

### 视觉效果示例

| 多模板预览 | 方法定位页 |
| --- | --- |
| ![Quad-LCD preview grid](skills/uav-paper-report/assets/screenshots/quad-lcd-preview-grid.jpg) | ![Quad-LCD analysis](skills/uav-paper-report/assets/screenshots/quad-lcd-analysis.jpg) |

| 强化学习方法图 | 实验结果表 | CBF 公式排版 |
| --- | --- | --- |
| ![RL method diagram](skills/uav-paper-report/assets/screenshots/rl-method-diagram.jpg) | ![RL result table](skills/uav-paper-report/assets/screenshots/rl-result-table.jpg) | ![CBF formula layout](skills/uav-paper-report/assets/screenshots/cbf-formula-layout.jpg) |

| FoV-CBF 示例 | PRIMER 示例 | SPOT 示例 |
| --- | --- | --- |
| ![FoV-CBF preview grid](skills/uav-paper-report/assets/screenshots/fov-cbf-preview-grid.png) | ![PRIMER preview grid](skills/uav-paper-report/assets/screenshots/primer-preview-grid.png) | ![SPOT preview grid](skills/uav-paper-report/assets/screenshots/spot-preview-grid.png) |

### 安装方式

把 skill 目录复制到 Codex 的 skills 目录下：

```powershell
Copy-Item -Recurse -Force .\skills\uav-paper-report $env:USERPROFILE\.codex\skills\uav-paper-report
```

安装后，在 Codex 中提出类似请求即可触发：

```text
读一篇无人机轨迹规划论文，按照我给的 PPT 模板生成中文论文汇报。
```

也可以直接指定这个 skill：

```text
使用 uav-paper-report skill，读取这篇论文，生成一份 20 页左右的中文组会汇报。
方法部分详细一些，实验部分简洁一些，公式需要配合解释。
```

### 运行环境

基础生成和检查通常需要：

- Python 3.10+
- `python-pptx`
- `Pillow`
- `PyMuPDF`
- LibreOffice，用于无界面导出 PDF

scaffold 的 Python 依赖可以参考：

```text
skills/uav-paper-report/assets/scaffolds/requirements.txt
```

### 推荐用法

更容易得到稳定结果的请求通常会包含三类信息：

- 论文：PDF、论文链接、标题，或者已经整理好的论文内容。
- 模板：历史 PPT、空模板、示例 PPT，或者明确的风格要求。
- 汇报偏好：页数、详细程度、重点章节、是否需要公式、是否需要多图、多表或多实验解释。

示例：

```text
使用这个模板生成一篇无人机论文汇报，20 页左右。
背景部分 3 页以内，方法部分写详细一些，需要公式和算法流程图。
实验结果用表格和图说明，不要只贴图。
```

```text
检查这份生成好的 PPT，重点看文字是否重叠、图片比例是否异常、
有没有无意义换行、每页空白是否太多，修正后再给我预览。
```

### 质量标准

这个 skill 对 PPT 的要求比较明确：

- 正文字体统一，默认使用 Times New Roman。
- 不同层级字号固定，不允许为了塞内容在每页随意变大变小。
- 正文段落要有项目符号和合理缩进，不能靠空段落制造间距。
- 重点可以加粗或标红，但只强调关键词和关键数字，不能整段变红。
- 公式尽量使用可编辑的 PPT 文本或 shape 组合，字号和对齐要统一。
- 图片必须保持比例，不裁掉关键信息，不和文字、表格、公式重叠。
- 每页都要有足够内容，普通内容页不应留下明显超过约 20% 的无意义空白。
- 表述要像正式论文汇报，只说明论文做了什么、结果说明什么、方法限制在哪里。
- 不写“讲解重点”“应该强调”这类提示给汇报人的话。

这些规则来自多轮实际生成和修改中反复出现的问题：字号漂移、无意义换行、图片重叠、公式难看、空白过多、模板首尾页颜色不一致等。

### 自动检查和修复

仓库里包含一些脚本，用来把人工检查中反复遇到的问题自动化：

- `audit_pptx_text.py`: 检查空文本框、手动换行、异常段落间距、字号漂移、越界 shape、文字和图表重叠等。
- `scan_rendered_slides.py`: 检查渲染后的 PNG 是否存在大面积空白、拥挤、内部空白带等问题。
- `render_pptx_previews.py`: 把导出的 PDF 渲染为逐页 PNG 和预览图。
- `repair_pptx_layout.py`: 发布前清理空文本体、统一部分字号漂移、移除正文里的手动换行。
- `run_template_matrix.py`: 对公开示例做多论文、多模板回归检查。
- `run_template_smoke.py`: 用本地模板批量生成压力测试页，检查模板泛化能力。
- `sanitize_pptx_privacy.py`: 清理 PPTX 元数据、备注、批注、自定义属性和可见个人信息。

这些脚本不是为了把 PPT 自动改到完美，而是帮助 Codex 更快发现明显问题。最终仍然需要查看渲染后的页面。

### 仓库内容

```text
paper2ppt/
├── README.md
├── LICENSE
└── skills/
    └── uav-paper-report/
        ├── SKILL.md
        ├── agents/openai.yaml
        ├── scripts/
        ├── references/
        └── assets/
            ├── examples/
            ├── scaffolds/
            ├── screenshots/
            └── template-profiles/
```

主要目录说明：

- `skills/uav-paper-report/SKILL.md`: Codex 使用该 skill 时读取的核心说明。
- `skills/uav-paper-report/references/`: 更细的风格、依赖和 QA 检查说明。
- `skills/uav-paper-report/scripts/`: PPT 审查、渲染、修复、隐私清理和多模板测试脚本。
- `skills/uav-paper-report/assets/examples/`: 已清理的 AI 生成示例 PPT。
- `skills/uav-paper-report/assets/scaffolds/`: 可复用的 `python-pptx` 生成脚本。
- `skills/uav-paper-report/assets/screenshots/`: README 中使用的效果截图。
- `skills/uav-paper-report/assets/template-profiles/`: 多模板测试和压力测试配置。

### 当前状态

目前仓库已经整理了 8 个公开 AI 生成示例 PPTX，覆盖轨迹规划、强化学习导航、多机 CBF 规划、覆盖路径规划、FoV-CBF、PRIMER、SPOT 等方向。示例和脚本已经围绕这些常见问题做了多轮修正：

- 文字和图片重叠
- 正文字号不统一
- 无意义换行和空段落
- 图片比例、裁剪和说明不稳定
- 公式区域过挤或过空
- 表格比例不协调
- 页面空白过多
- 模板首尾页风格不一致
- 公开文件中的个人信息残留

这些示例不是最终产品，而是给后续生成任务提供可复用的质量基线。

### 后续目标

- 更自动地抽取用户 PPT 模板的版式、配色、字号层级和常用页面结构。
- 增加更明确的内容密度参数，让背景、方法、实验、总结可以分别控制详略。
- 改进 PowerPoint 原生公式生成，让复杂公式更接近正式 PPT 的书写方式。
- 加强图片裁剪检测，减少子图被截断、caption 过多、图片比例异常等问题。
- 扩展到更多机器人和自主系统方向，如 SLAM、控制、多智能体协同和机器人学习。
- 封装更完整的命令行流程：输入论文 PDF、模板 PPTX 和配置文件，输出 PPTX、PDF、PNG 预览和 QA 报告。

### 隐私说明

公开仓库不应包含私人模板、真实汇报人姓名、导师信息、课题组内部材料、备注页、批注作者或文档元数据。发布前建议运行：

```powershell
python .\skills\uav-paper-report\scripts\sanitize_pptx_privacy.py --check-only --fail-on-warning .\skills\uav-paper-report\assets\examples
```

## English

`paper2ppt` contains a Codex Skill named `uav-paper-report`.

It is not a finished one-click PDF-to-PPT product. It is a reusable generation and review workflow that helps Codex read a paper, understand a presentation template, organize a Chinese academic report, place formulas and figures, and visually check the final deck.

The current skill is focused on Chinese paper-report decks for UAV, robotics, trajectory planning, SLAM, control, and autonomous-systems papers. The same workflow can also be adapted to other academic presentation tasks.

### What Problem It Solves

A useful paper-report deck is more than a compressed abstract. In practice, the hard parts are:

- Extracting a talkable story from the paper: motivation, problem, method, formulas, experiments, and limitations.
- Reusing the user's template style: cover, ending slide, headers, rules, colors, font hierarchy, tables, and formula regions.
- Making method slides dense enough to explain, but not so crowded that formulas and text collide.
- Cropping figures cleanly without stretching, truncating, or overlapping nearby text.
- Keeping bullets, indentation, font size, red emphasis, and spacing consistent across the deck.
- Rendering the PPTX and checking the actual pages, not just the file structure.

`uav-paper-report` turns these requirements into reusable instructions, scripts, and examples.

### When To Use It

Use this skill when you want to:

- Generate a Chinese group-meeting or thesis-style paper report from a UAV or robotics paper.
- Match a user-provided PPT template or an existing report style.
- Improve a generated deck until it looks closer to a human-made academic report.
- Audit a PPT for overlap, bad image proportions, excessive blank space, meaningless line breaks, and inconsistent font sizes.
- Test generation across multiple papers and templates.
- Sanitize public example decks before publishing.

This repository is best understood as a Codex skill package and workflow reference, not a hosted web application.

### What It Generates

The default target is an editable Chinese report deck with about 18-24 slides:

- Cover and paper information
- Research background and motivation
- Related work or gaps in existing methods
- Main idea and technical route
- Method details, equations, and algorithm flow
- Experimental setup, metrics, and result tables
- Ablations, comparisons, or visual results
- Limitations, applicability, and summary
- A closing slide that matches the template style

The content density can be adjusted. For example:

- Make the method section more detailed and keep the background short.
- Add more formulas while compressing experimental details.
- Keep the report under 15 slides.
- Build a 25-slide version for a 30-minute talk.
- Spend more pages on experiments and fewer pages on related work.

The goal is not mechanical expansion. Each slide should have a clear claim, evidence, and interpretation.

### Template Adaptation

When a template or historical deck is available, the workflow is:

1. Inspect slide size, cover, ending slide, headers, rules, and palette.
2. Identify font levels for titles, main bullets, sub-bullets, tables, notes, and formulas.
3. Decide which layout families the template supports: vertical stack, left-right comparison, figure plus explanation, table plus interpretation, or formula band.
4. Choose slide layouts from the paper content instead of forcing every slide into one pattern.
5. Export the generated PPTX to PDF, render slide PNGs, and inspect the actual output for overlap, overflow, and excessive blank space.

The public repository does not include private templates or personal report materials. It only includes AI-generated examples, screenshots, scripts, and public configuration files.

### Example Outputs

| Multi-template preview | Method positioning slide |
| --- | --- |
| ![Quad-LCD preview grid](skills/uav-paper-report/assets/screenshots/quad-lcd-preview-grid.jpg) | ![Quad-LCD analysis](skills/uav-paper-report/assets/screenshots/quad-lcd-analysis.jpg) |

| RL method diagram | RL result table | CBF formula layout |
| --- | --- | --- |
| ![RL method diagram](skills/uav-paper-report/assets/screenshots/rl-method-diagram.jpg) | ![RL result table](skills/uav-paper-report/assets/screenshots/rl-result-table.jpg) | ![CBF formula layout](skills/uav-paper-report/assets/screenshots/cbf-formula-layout.jpg) |

| FoV-CBF example | PRIMER example | SPOT example |
| --- | --- | --- |
| ![FoV-CBF preview grid](skills/uav-paper-report/assets/screenshots/fov-cbf-preview-grid.png) | ![PRIMER preview grid](skills/uav-paper-report/assets/screenshots/primer-preview-grid.png) | ![SPOT preview grid](skills/uav-paper-report/assets/screenshots/spot-preview-grid.png) |

### Installation

Copy the skill folder into your Codex skills directory:

```powershell
Copy-Item -Recurse -Force .\skills\uav-paper-report $env:USERPROFILE\.codex\skills\uav-paper-report
```

Then ask Codex for a task like:

```text
Read this UAV trajectory-planning paper and generate a Chinese paper-report PPT using my template.
```

You can also call the skill explicitly:

```text
Use the uav-paper-report skill to read this paper and create a 20-slide Chinese group-meeting report.
Make the method section detailed, keep the experiments concise, and include formulas with explanations.
```

### Runtime

Basic generation and review usually need:

- Python 3.10+
- `python-pptx`
- `Pillow`
- `PyMuPDF`
- LibreOffice for headless PDF export

The scaffold dependency file is:

```text
skills/uav-paper-report/assets/scaffolds/requirements.txt
```

### Recommended Prompts

Good requests usually include three things:

- Paper: a PDF, paper link, title, or extracted paper notes.
- Template: a PPTX template, a historical deck, or concrete style requirements.
- Preferences: slide count, detail level, important sections, formulas, figures, tables, and experiment depth.

Examples:

```text
Use this template to generate a UAV paper report with about 20 slides.
Keep the background within 3 slides. Make the method section detailed with formulas and an algorithm diagram.
Explain experimental figures and tables instead of only pasting them.
```

```text
Audit this generated PPT. Focus on text overlap, image proportions, meaningless line breaks,
excessive blank space, and inconsistent font sizes. Fix the deck and give me previews.
```

### Quality Expectations

The skill is opinionated about presentation quality:

- Body text should use a consistent font, currently Times New Roman by default.
- Font levels should stay fixed across the deck instead of changing slide by slide.
- Body paragraphs should use bullets and proper indentation.
- Empty paragraphs and manual body line breaks should not be used for spacing.
- Red or bold emphasis should stay on key terms and important numbers, not whole paragraphs.
- Formulas should be editable PowerPoint text or shape-based components when possible.
- Images must preserve aspect ratio, keep important content visible, and avoid overlap.
- Normal content slides should not leave large meaningless blank regions.
- Wording should sound like a real academic report: state what the paper does, what the results show, and where the method is limited.
- Avoid presenter-advice phrases such as "key point to explain", "should emphasize", or "report wording".

These rules come from repeated real corrections around font drift, line breaks, image overlap, weak formula layout, excessive blank space, and inconsistent template styling.

### Automation Helpers

The repository includes scripts for recurring checks:

- `audit_pptx_text.py`: checks empty text bodies, manual newlines, abnormal spacing, font drift, out-of-bounds shapes, and content overlap.
- `scan_rendered_slides.py`: checks rendered PNGs for large blank areas, crowding, and internal whitespace bands.
- `render_pptx_previews.py`: renders exported PDFs into slide PNGs and preview grids.
- `repair_pptx_layout.py`: cleans empty text bodies, some font drift, and body newlines before publishing.
- `run_template_matrix.py`: runs public multi-paper, multi-template regression checks.
- `run_template_smoke.py`: generates stress decks across local PPTX templates to test generalization.
- `sanitize_pptx_privacy.py`: removes personal metadata, notes, comments, custom properties, and visible private information.

These scripts are quality helpers. The final deck still needs visual inspection after rendering.

### Repository Layout

```text
paper2ppt/
├── README.md
├── LICENSE
└── skills/
    └── uav-paper-report/
        ├── SKILL.md
        ├── agents/openai.yaml
        ├── scripts/
        ├── references/
        └── assets/
            ├── examples/
            ├── scaffolds/
            ├── screenshots/
            └── template-profiles/
```

Key paths:

- `skills/uav-paper-report/SKILL.md`: core instructions loaded by Codex.
- `skills/uav-paper-report/references/`: detailed style, dependency, and QA notes.
- `skills/uav-paper-report/scripts/`: audit, render, repair, privacy, and multi-template scripts.
- `skills/uav-paper-report/assets/examples/`: sanitized AI-generated example PPT decks.
- `skills/uav-paper-report/assets/scaffolds/`: reusable `python-pptx` generation scripts.
- `skills/uav-paper-report/assets/screenshots/`: screenshots used in this README.
- `skills/uav-paper-report/assets/template-profiles/`: multi-template test configuration.

### Current Status

The repository currently includes 8 public AI-generated example PPTX decks covering trajectory planning, RL navigation, multi-quadrotor CBF planning, coverage path planning, FoV-CBF, PRIMER, SPOT, and related UAV/autonomy topics.

The examples and scripts have been iterated around issues that repeatedly appeared in real use:

- Text and image overlap
- Inconsistent body font sizes
- Meaningless line breaks and empty paragraphs
- Unstable image crop and aspect ratio
- Formula regions that are too crowded or too sparse
- Poor table proportions
- Excessive blank space
- Inconsistent cover and ending-slide styling
- Personal information left in public files

These examples are quality baselines for future generation tasks, not final product demos.

### Roadmap

- Extract template style more automatically: layout families, palette, font hierarchy, table style, and common slide structures.
- Add clearer content-density controls for background, method, experiments, and summary sections.
- Improve native PowerPoint formula generation for more standard academic equation layout.
- Strengthen figure-crop QA for truncated subfigures, caption-heavy crops, and aspect-ratio errors.
- Expand examples to more robotics and autonomous-systems topics, including SLAM, control, multi-agent coordination, and robot learning.
- Package a fuller CLI workflow: input PDF, template PPTX, and config file; output PPTX, PDF, PNG previews, and QA report.

### Privacy

Public assets should not contain private templates, real presenter names, advisor information, internal group materials, notes slides, comment authors, or document metadata. Before publishing, run:

```powershell
python .\skills\uav-paper-report\scripts\sanitize_pptx_privacy.py --check-only --fail-on-warning .\skills\uav-paper-report\assets\examples
```
