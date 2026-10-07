# Runbook: 模板 + PPT → 前四项

所有命令都在工作目录（`work/`、`outputs/` 所在处）执行。`<office>` 指
`~/.codex/skills/office-com-windows/scripts`，`<figs>` 指
`~/.codex/skills/pptx-figure-extract/scripts`，`<self>` 指本 skill 的 `scripts`。
`py` 指内置 Python：
`C:\Users\<user>\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe`。

## 0. 准备

把用户给的两个文件复制进 `work/`，改成 ASCII 名字（`template_src.doc`、`ppt_src.pptx`）。
中文路径在后面的 PowerShell COM 脚本里是麻烦的来源，复制一次就一劳永逸。

```
work/template_src.doc      ← 实验报告模板
work/ppt_src.pptx          ← 实验 PPT
work/template.docx         ← 转出来的可编辑模板
work/ppt_dump/             ← PPT 文字与媒体索引
work/slides_hi/            ← 幻灯片渲染图
outputs/report_assets/     ← 交付的图片
```

## 1. 读模板

### 1.1 `.doc` → `.docx`（Flat OPC 路线）
不要用 Word 的 SaveAs：目标路径在用户目录下会永久卡死。

```powershell
$env:WORD_SRC="…\work\template_src.doc"; $env:WORD_OUT="…\work\template_flat.xml"; $env:WORD_MODE="flatopc"
Start-Process powershell -ArgumentList "-NoProfile","-ExecutionPolicy","Bypass","-File","<office>\word_read.ps1" -WindowStyle Hidden -Wait
```
```
python <office>\flat2docx.py work\template_flat.xml work\template.docx
```
转换后用 python-docx 自查一次：`Document(...)` 能打开、段落数和表格数跟原文档一致。

### 1.2 摸清结构
```powershell
$env:WORD_SRC="…\work\template.docx"; $env:WORD_OUT="…\work\template_dump.txt"; $env:WORD_MODE="structure"
$env:WORD_HEADS="一、实验目的|实验原理|实验仪器|四、实验内容与步骤|五、数据处理"
Start-Process powershell -ArgumentList "-NoProfile","-ExecutionPolicy","Bypass","-File","<office>\word_read.ps1" -WindowStyle Hidden -Wait
```

要确认的三件事：

1. **前四项到底是哪四项** —— 看段落文本 + 表格行。深大模板的正文是一个 9 行单列表格，
   前四行依次是 `一、实验目的`、`实验原理`、`实验仪器`、`四、实验内容与步骤`。
2. **编号方式** —— dump 里的 `num='二、'` 说明该标题靠 `numbering.xml` 自动编号
   （`chineseCounting`、起始值 2）。写正文时绝不能给段落加 `w:numPr`。
3. **版心** —— `page W H … left right` 给出页边距；图宽上限取"正文栏宽"再留 1–2 cm
   余量（深大模板正文栏宽约 14.6 cm，插图按 ≤11.5 cm 处理就安全）。

## 2. 读 PPT

```bash
python <figs>\pptx_figures.py dump work\ppt_src.pptx work\ppt_dump
python <figs>\pptx_figures.py media work\ppt_src.pptx work\media
```
```powershell
$env:PPT_SRC="…\work\ppt_src.pptx"; $env:PPT_OUT="…\work\ppt_com.txt"; $env:PPT_PNG_DIR="…\work\slides_hi"
$env:PPT_PNG_W="2560"; $env:PPT_PNG_H="1920"
Start-Process powershell -ArgumentList "-NoProfile","-ExecutionPolicy","Bypass","-File","<office>\pptx_dump.ps1" -WindowStyle Hidden -Wait
```

`ppt_com.txt` 里带上备注页（`notes:`）和表格；`slides_hi/slide-NN.png` 用来肉眼确认每张图画的
到底是什么、有没有箭头/圈注、有没有水印。

`.ppt`（老格式）先复制到 `C:\CodexTmp`，用 PowerPoint COM 另存为 `.pptx`（`ppSaveAsOpenXMLPresentation = 24`）再走上面的流程。

## 3. 建立映射表

动笔前先写出这张表，写进工作笔记，也是最后回答用户"哪来的"的依据：

| 模板项 | PPT 页 | 要写的内容 | 用图 | 用公式 |
| --- | --- | --- | --- | --- |
| 一、实验目的 | 4 | 3 条目的，原文 | — | — |
| 实验原理 | 2,3,5,6,7,8 | 光谱产生、光栅色散、分辨本领与角色散、结构与光路 | 图2-1…2-3 | 式(1)…(5) |
| 实验仪器 | 9,10 | 光学/电子/软件三部分，要点式 | 图3-1 | — |
| 四、实验内容与步骤 | 11–16 | 6 个步骤，①②③ 条目式 | 图4-1、4-2 | — |

## 4. 取图

按 `pptx-figure-extract` 的规则选图，写 `figures.json`：

```json
[
 {"out":"ppt_p02_img1.png","media":"image3.png","slide":2,"desc":"氢原子能级跃迁图"},
 {"out":"ppt_p12_img14.png","slide":12,"box":[611188,3992563,8162925,2074862],"desc":"开机提示对话框"}
]
```
```bash
python <figs>\pptx_figures.py build work\ppt_src.pptx --slides-png work\slides_hi \
    --media-dir work\media --spec work\figures.json --out outputs\report_assets
```
清单自动写成 `图片清单.txt`；后来精简掉的图把"用途"列改成"未采用（精简后未插入）"，
这样用户想补回时找得到。

## 5. 写 content.py

照 `scripts/docx_report.py` 头部注释的块格式写。要点：

* 段落文字用 PPT 原话，长句可以断句、合并，但不改术语、参数、单位、条件。
* 独立公式写 `("eq", 式子, 编号)`，式子用 `omml` 的 `V()/P()/frac()/sub()/…` 拼。
* 行内符号用普通 Unicode 文字（`λ`、`Δλ`、`M₁`、`S₁`、`d = a + b`），不要写成 LaTeX。
* 图注固定格式：`图 {章}-{序}  标题（来源：PPT 第 N 页）`。

## 6. 组装

```bash
python <self>\docx_report.py --template work\template.docx --spec work\content.py \
    --assets outputs\report_assets --out "outputs\实验报告_前四项_完成版.docx"
```
脚本会：填封面"实验名称"（若 spec 给了 `EXPERIMENT_NAME`）、清掉前四行里的占位空行、
按块写入、把"图+图注"包成不可分页的块、最后保存为新文件（拒绝覆盖模板）。

## 7. 验收

```powershell
$env:WORD_SRC="…\outputs\实验报告_前四项_完成版.docx"; $env:WORD_OUT="C:\CodexTmp\report.pdf"
Start-Process powershell -ArgumentList "-NoProfile","-ExecutionPolicy","Bypass","-File","<office>\word_export_pdf.ps1" -WindowStyle Hidden -Wait
```
```bash
python <office>\pdf2png.py C:\CodexTmp\report.pdf C:\CodexTmp\report_png 150
python <office>\qa_layout.py C:\CodexTmp\report.pdf
```
```powershell
$env:DOCX_PATH="…\outputs\实验报告_前四项_完成版.docx"; $env:DOCX_OUT="C:\CodexTmp\inspect.txt"
$env:DOCX_HEADS="一、实验目的|实验原理|实验仪器|四、实验内容与步骤|五、数据处理|六、结果陈述|七、思考题"
Start-Process powershell -ArgumentList "-NoProfile","-ExecutionPolicy","Bypass","-File","<office>\docx_inspect.ps1" -WindowStyle Hidden -Wait
```

判据：

* **逐页看图**，确认无重叠、无裁切、无缺字、无越界。
* `qa_layout.py` 报出的"orphaned captions"必须是 none。
* `docx_inspect.ps1` 里第二、三项的 `num` 必须显示为"二、""三、"，`wide items=0`。
* 第五项及之后在渲染图里仍是空白模板。
* 底部空白 20–30% 属于正常分页（大图放不下会整块下移）；要求更紧凑时优先缩小插图上限，
  而不是砍正文。

## 8. 交付话术

回复里给出：完成清单、图片目录与清单文件、**被精简掉的 PPT 内容清单**、
待确认项（PPT 自相矛盾/被裁切/水印图/未填个人信息）、以及 2–3 个后续动作。
