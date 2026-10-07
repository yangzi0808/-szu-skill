# -szu-skill

用老师的实验 PPT 流程化完成深圳大学物理实验的**预实验报告**。

把一份「实验报告模板 + 实验 PPT」交给 Codex，它会读模板识别前四项的结构，逐页读 PPT
取插图与公式，把 **实验目的 / 实验原理 / 实验仪器 / 四、实验内容与步骤** 写进模板副本，
输出一份可直接打印的 Word 报告，以及一份标注了来源页码的图片素材目录。

Skill for writing Shenzhen University College Physics pre-lab reports from the
instructor's slides.

## 解决什么问题

预实验报告每次的动作完全一样，手工做却要一两个小时：

- PPT 几十页，原理图、装置图、操作照片散落在各页，还得挑掉 Logo、水印和重复截图；
- 公式在 PPT 里是图片，抄进 Word 要么糊要么变形；
- 报告模板是固定格式的表格，字号、编号、图注格式一乱就得返工；
- 很容易写成“PPT 搬家”，报告冗长，重点反而不突出。

这个仓库把这一套动作固化成流水线：**读模板 → 读 PPT → 选图取图 → 写内容 → 排版 → 渲染验收**。

## 它能做到什么

| 方面 | 说明 |
| --- | --- |
| 结构 | 只填模板的前四项，第五项（数据处理）及之后保持空白；模板原件绝不覆盖 |
| 插图 | 从 PPT 抽取原始图片，或从幻灯片高清渲染图上裁切（含箭头、圈注的那种），并按规则剔除水印图、装饰图和重复截图 |
| 命名与溯源 | 统一命名 `ppt_p{页码}_img{序号}.png`，每张图带图注 `图 2-1 标题（来源：PPT 第 2 页）` |
| 公式 | 排成 **Word 原生公式**（矢量不失真、双击可编辑），独立公式居中并带右侧编号（1）（2）… |
| 精简 | 按参考报告的尺度浓缩，而不是把 PPT 原文誊抄一遍 |
| 验收 | 自动导出 PDF 并逐页渲染检查：图注被分页甩掉、表格被劈开、页面大片空白 |
| 交付说明 | 交付时列出被精简掉的 PPT 内容、需要用户确认的冲突项和未填字段 |

## 三个 skill 的分工

| Skill | 职责 | 单独使用的场景 |
| --- | --- | --- |
| **lab-report-first-four-sections** | 主流程：识别模板前四项、建 PPT 映射、生成报告 | 有模板 + PPT，要交预实验报告 |
| **pptx-figure-extract** | 取图：抽媒体、从渲染图裁切、生成图片清单 | 只想把 PPT 里的图拿出来复用 |
| **office-com-windows** | 底层 I/O：读 `.doc`/`.docx`、Flat OPC 转换、导出 PDF、渲染与版式校验 | 任何 Word/PPT 读取、导出、渲染校验的任务 |

主 skill 的 `SKILL.md` 会指名调用另外两个 skill 的脚本，三个一起装才能完整跑通。

## 快速开始

### 1. 安装

把三个目录复制到 Codex 的 skill 目录：

```powershell
$dst = Join-Path $env:USERPROFILE ".codex\skills"     # 若设置了 CODEX_HOME 则用 $env:CODEX_HOME\skills
Copy-Item -Path .\lab-report-first-four-sections, .\office-com-windows, .\pptx-figure-extract `
          -Destination $dst -Recurse -Force
```

### 2. 调用

在 Codex 里附上模板和 PPT，直接说：

```
实验报告模板和实验 PPT 见附件。请用 $lab-report-first-four-sections 完成前四项，
PPT 是唯一事实来源，输出到 outputs，并告诉我哪些内容被精简掉了。
```

也可以不点名，直接描述需求（“这是实验报告模板和实验 PPT，帮我完成前四项”），
Codex 会根据 skill 描述自动匹配。只想取图时用
`$pptx-figure-extract 把这个 PPT 里的原理图和装置图都提出来`。

### 3. 产出

```
outputs/实验报告_前四项_完成版.docx      # 可直接打印/提交
outputs/report_assets/
    ppt_p02_img1.png  …                 # 按 PPT 页码命名的插图
    图片清单.txt                         # 文件名 / 尺寸 / 来源 / 内容 / 用途
```

## 工作流程

1. **读模板** —— 把 `.doc` 转成可编辑的 `.docx`（走 Flat OPC，不依赖 Word 的另存为），
   读出段落、样式、自动编号和表格，确认前四项到底是哪四项。
2. **读 PPT** —— 逐页导出文字、表格、备注页和每页高清图；同时导出 `ppt/media/` 下所有原始图片。
3. **建映射** —— 模板项 ↔ PPT 页码 ↔ 要写的文字 ↔ 用哪张图 ↔ 用哪个公式，先列成表再动笔。
4. **选图取图** —— 对比图形与源图的宽高比判断 PPT 是否裁切过，决定“直接复制原图”还是
   “从渲染图裁切”；剔除水印图与装饰图。
5. **写内容** —— 用 PPT 原术语、原参数、原单位，按预实验报告的体量浓缩。
6. **排版组装** —— 文字、插图（含图注）、原生公式、小表格，按模板格式写进副本。
7. **渲染验收** —— 导出 PDF 逐页看图，检查图注、分页、空白和自动编号，通过后才交付。

## 环境要求

- **Windows**：底层脚本通过 Word / PowerPoint COM 驱动，仅 Windows 可用。
- **Microsoft Office 16**（Word + PowerPoint）：读取 `.doc`、导出 PDF、渲染幻灯片、版式校验都要用。
- **不需要 LibreOffice**：渲染校验走 Word 导出 PDF 再转 PNG 的路线，不依赖 LibreOffice。
- **Python 3.9+**，依赖：`python-docx`、`python-pptx`、`Pillow`、`lxml`、`numpy`、`pdfplumber`、`pypdfium2`。

## 目录结构

```
.
├── lab-report-first-four-sections/      # 主 skill
│   ├── SKILL.md                         流程、铁律、交付要求
│   ├── agents/openai.yaml               UI 元数据
│   ├── references/
│   │   ├── workflow.md                  完整操作手册（含各脚本命令）
│   │   ├── content-conventions.md       写什么、写多少、怎么精简
│   │   └── equations.md                 公式排版约定与示例
│   └── scripts/
│       ├── docx_report.py               组装库：按 spec 填模板、插图注、表格、公式
│       └── omml.py                      Word 原生公式（OMML）构造器
├── office-com-windows/                  # Office 自动化工具箱
│   ├── SKILL.md
│   ├── references/environment.md        环境事实与排障记录
│   └── scripts/
│       ├── word_read.ps1                读 .doc/.docx，或导出 Flat OPC
│       ├── flat2docx.py                 Flat OPC → 真正的 .docx
│       ├── word_export_pdf.ps1          导出 PDF（必须写到用户目录之外）
│       ├── pdf2png.py                   PDF → 逐页 PNG
│       ├── qa_layout.py                 底部空白 + 图注错页检查
│       ├── docx_inspect.ps1             页数、自动编号、图片尺寸核查
│       └── pptx_dump.ps1                幻灯片文字/表格/备注 + 导出 PNG
└── pptx-figure-extract/                 # 取图
    ├── SKILL.md
    ├── references/figure-selection.md   哪些图该留、哪些该丢
    └── scripts/pptx_figures.py          dump / media / build 三个子命令
```

## 注意事项

- **Word 只能写到用户目录之外。** 在本机验证过的行为是：`SaveAs2` / `ExportAsFixedFormat`
  一旦把目标路径放在 `C:\Users\<用户>\…`（文档、桌面、`%TEMP%`）下就会永久卡死，写到
  `C:\CodexTmp` 之类的路径则秒回（疑似杀毒软件的过滤驱动所致）。所以报告用 python-docx
  生成，渲染用的 PDF 一律先导出到用户目录之外——照着做没有副作用。
  细节见 `office-com-windows/references/environment.md`。
- **PowerShell 脚本保持 ASCII。** Windows PowerShell 5.1 以 ANSI 读取 `.ps1`，脚本里出现中文
  会变成乱码甚至解析失败，所以中文路径通过环境变量传入。
- **不要用 Word 的“另存为”做 `.doc` → `.docx`**，走 `word_read.ps1`（Flat OPC）+ `flat2docx.py`。
- **在别的机器上使用**：`lab-report-first-four-sections/SKILL.md` 里写了本机的 skill 绝对路径，
  换机器时改成自己的 skill 目录即可，其余脚本用的是环境变量和 `$PSScriptRoot`，不受影响。

## 已知限制

- 目前只覆盖**预实验报告的前四项**；数据处理、结果陈述、思考题（需要原始数据）尚未纳入。
- 老格式 `.ppt` 需要先转成 `.pptx` 再取图（`office-com-windows` 可以读 `.ppt` 并导出幻灯片图）。
- 整页说明书的扫描图不会被当成插图；其文字会被转写进正文。
- 分辨率过低的源图（例如 200×153 的截图）不会硬塞进报告，会在交付说明里单独提示。

## 许可证

[MIT License](LICENSE) © 2026 yangzi0808
