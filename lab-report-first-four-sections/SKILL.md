---
name: lab-report-first-four-sections
description: Complete the first four sections (一、实验目的 / 实验原理 / 实验仪器 / 四、实验内容与步骤) of a Chinese university lab-report template from an experiment PPT, producing a shortened, faithfully sourced .docx with extracted figures, numbered native Word equations and a source page on every caption. Use when the user supplies a lab-report template plus an experiment PPT/幻灯片 and asks to write or fill in 前四项, or asks to condense such a report.
metadata:
  short-description: 用 PPT 完成实验报告模板前四项并排版公式与插图
---

# 实验报告前四项

输入 = 一份实验报告模板（.doc/.docx）+ 一份实验 PPT（.ppt/.pptx）。
输出 = 模板副本中把前四项写好，配图、配公式，其余部分保持原样。

## 铁律

1. **PPT 是唯一事实来源。** 公式、参数、单位、步骤、结论都照 PPT 写，不用网络资料，不凭常识补内容。PPT 里没有的写【PPT未说明，需补充】。
2. **模板只决定结构和格式。** 标题层级、编号方式、字体字号、表格框架都照模板；内容冲突时以 PPT 为准。
3. **只做前四项。** 先读模板确认前四项到底是哪四项（常见是 一、实验目的 / 实验原理 / 实验仪器 / 四、实验内容与步骤，其中第二、三项有时靠模板的自动编号显示为"二、""三、"）。第五项及之后一律不填、不改。
4. **不覆盖模板原件。** 输出新文件（建议 `outputs/实验报告_前四项_完成版.docx`）。
5. **要精简，不要照抄。** 报告是 PPT 的浓缩版，不是 PPT 的搬家。详见
   [references/content-conventions.md](references/content-conventions.md)。
6. **公式排版成 Word 原生公式**（可编辑、打印清晰），不要留 LaTeX 源码，也不要用公式截图。详见
   [references/equations.md](references/equations.md)。
7. **关键信息缺失就先问。** 前四项识别不出来，或 PPT 缺完成前四项所需的关键参数时，列出缺失项再问用户，不要猜。

## 依赖的两个 skill

文档 I/O 和图片提取都在别的 skill 里，不要重复实现：

| 需要的动作 | 用哪个 skill | 脚本位置 |
| --- | --- | --- |
| 读 .doc/.docx、Flat OPC 转 docx、导出 PDF、渲染校验 | `office-com-windows` | `C:\Users\杨千贤\.codex\skills\office-com-windows\scripts\` |
| 读 .ppt/.pptx、导出幻灯片 PNG、提取图片 | `pptx-figure-extract` | `C:\Users\杨千贤\.codex\skills\pptx-figure-extract\scripts\` |

本 skill 自己的脚本在
`C:\Users\杨千贤\.codex\skills\lab-report-first-four-sections\scripts\`。

**动手前先读 `office-com-windows/SKILL.md`。** 里面有两条会让你白花半小时的坑：
Word 只能写到用户目录之外（如 `C:\CodexTmp`），以及必须用 `/a` 启动 Word 才能自动化。

## 流程

完整命令见 [references/workflow.md](references/workflow.md)。骨架是：

1. **读模板**：把用户的 .doc 先复制成 ASCII 文件名放进 `work/`，用
   `word_read.ps1`（`WORD_MODE=flatopc`）+ `flat2docx.py` 得到可编辑的 .docx，再用
   `WORD_MODE=structure` 读出段落、样式、自动编号和表格，确认前四项标题与结构。
2. **读 PPT**：`pptx_figures.py dump` 出逐页文字/表格/备注、图形与媒体的对应、媒体清单；
   用 `pptx_dump.ps1` 以 2560×1920 导出每页 PNG。（.ppt 先复制再读，别碰用户打开的文件。）
3. **建映射**：模板项 → PPT 页码 → 该写什么 → 用哪张图 → 哪些公式。列成表再动笔。
4. **选图与提取**：按 `pptx-figure-extract` 的规则筛掉水印图、装饰图、重复截图，用
   `pptx_figures.py build` 生成 `outputs/report_assets/ppt_p{页}_img{序号}.png` 和清单。
5. **写内容**：按 `references/content-conventions.md` 的尺度和写法把一段段文字写成
   `content.py`（spec 格式见 `scripts/docx_report.py` 头部注释）。
6. **组装**：`python scripts/docx_report.py --template … --spec … --assets … --out …`。
7. **渲染验收**：导出 PDF → 渲染成页图 → **逐页看图**，再跑 `qa_layout.py` 查底部空白和
   被分页甩掉的图注，跑 `docx_inspect.ps1` 确认页数、自动编号（"二、""三、"是否正常）
   和图片尺寸。有图注被拆、表格被劈开、大片空白就要回去调，不要直接交付。

## 交付

```
outputs/实验报告_前四项_完成版.docx
outputs/report_assets/ppt_p{页}_img{序号}.png     # 从 PPT 提取的图
outputs/report_assets/图片清单.txt                 # 文件名/尺寸/来源/内容/用途
```

回复用户时说清四件事：写了哪几项、图放在哪、**哪些 PPT 内容被精简掉了**（对方可能要补回）、
以及需要确认的项（PPT 内部矛盾、被裁切的文字、被跳过的水印图、模板里没填的个人信息字段）。

用 `:codex-followup[...]{prompt="..."}` 给 2–3 个后续动作，例如补回截图、再精简、导出 PDF。

## 已知的常客问题

- 模板第二、三项靠自动编号显示成"二、""三、"。**不要**给正文段落加 `w:numPr`，否则编号会串。
- 模板正文常是一个 9 行单列表格，前四行是前四项。写完 5 行以后的内容要保持空。
- 图片和它的图注会被分页拆开（表格单元格里 `keepNext` 无效），所以 `docx_report.py`
  把"图 + 图注"包成一个无框 1×1 嵌套表并设 `cantSplit`。不要退回用 `keepNext`。
- PPT 里常有笔误（同音字、上下标写反、图注与正文矛盾）。照原文保留明显口径，把不一致写进
  交付说明让用户定夺；纯同音错别字（如"纳灯"→"钠灯"）可以直接更正并说明。
