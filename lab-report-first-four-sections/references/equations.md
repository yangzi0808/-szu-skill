# 在 Word 里排公式：原生 OMML

## 为什么不用别的办法

| 办法 | 结论 |
| --- | --- |
| 把 LaTeX 源码当文字写进正文 | 用户明确否决过：打印出来像没排版的代码 |
| 公式截图 | 打印会糊，且不可编辑；用户也否决过"只放公式截图" |
| MathJax 渲染成 PNG | 本机没有 `mathjax-full`（Node 只有 `sharp`），Python 也没有 matplotlib，装依赖不划算 |
| **Word 原生公式（OMML）** | 矢量、任意缩放都清晰、双击可编辑，Word 自己排版，采用这个 |

效果对标：跟参考报告一样——独立公式居中带右侧编号，行内符号是普通文字。

## 两个约定

1. **独立公式**：居中 + 右侧编号 `（1）（2）…`
2. **行内符号**：普通 Unicode 文字。`λ`、`Δλ`、`Δθ`、`ν`、`M₁`、`S₂`、`d = a + b`。
   文档里不允许出现 `$`、`\frac`、`\lambda` 之类的痕迹。

## 独立公式怎么搭

`scripts/omml.py` 提供最小构造器，全部返回元素列表，用 `+` 拼接：

| 函数 | 作用 |
| --- | --- |
| `V("N")` | 斜体变量（数学变量） |
| `P("=")` | 正体（数字、运算符、单位、函数名） |
| `frac(num, den)` | 分式 |
| `sub(base, sub)` / `subsup(base, sub, sup)` | 下标 / 上下标 |
| `delim([...], "[", "]")` | 自适应括号 |
| `func("sin", V("θ"))` | 正体函数名 + 参数 |

放进文档由 `docx_report.py` 的 `("eq", 式子, 编号)` 块完成：它把 `m:oMath` 塞进一个
带"居中制表位 + 右制表位"的普通段落，Word 就会把公式居中、把编号顶到右边——这正是 Word
自己排编号公式的做法。制表位位置取单元格正文宽度（`cell_text_twips`）。

## 本项目的五个公式（可当模板抄）

```python
from omml import V, P, frac, sub, subsup, delim, func

# ΔE = hν = 13.6[1/n1² − 1/n2²] eV
eq1 = (P("Δ") + V("E") + P("=") + V("h") + V("ν") + P("=") + P("13.6")
       + delim([frac(P("1"), subsup(V("n"), P("1"), P("2"))), P("−"),
                frac(P("1"), subsup(V("n"), P("2"), P("2")))], "[", "]")
       + P("eV"))

# d = 1/N
eq2 = V("d") + P("=") + frac(P("1"), V("N"))

# d sinθ = kλ,  k = 0, ±1, ±2…
eq3 = (V("d") + func("sin", V("θ")) + P("=") + V("k") + V("λ") + P(",  ")
       + V("k") + P("=") + P("0") + P(",") + P("±") + P("1") + P(",")
       + P("±") + P("2") + P("…"))

# R = λ/Δλ = KN
eq4 = (V("R") + P("=") + frac(V("λ"), P("Δ") + V("λ")) + P("=")
       + V("K") + V("N"))

# D = Δθ/Δλ = k/(d cosθ)
eq5 = (V("D") + P("=") + frac(P("Δ") + V("θ"), P("Δ") + V("λ")) + P("=")
       + frac(V("k"), V("d") + func("cos", V("θ"))))
```

在 spec 里对应：
```python
("eq", eq1, 1),
("eq", eq2, 2),
```

## 自查清单

- [ ] 文档里搜不到 `$`、`\frac`、`\lambda`（用 `docx` 解包 `word/document.xml` 搜）。
- [ ] `word/document.xml` 里 `<m:oMath>` 的数量 == 公式数量，根节点声明了 `xmlns:m`。
- [ ] 渲染出的 PDF 里分式是上下叠的、下标真的在右下角、编号在右侧对齐。

## 常见坑

* `docx.oxml.OxmlElement("m:r")` 依赖 `docx.oxml.ns.nsmap` 里的 `m` 前缀；python-docx
  自带，写完后根节点会自动声明 `xmlns:m`，不用手工加。
* `m:r` 里子元素顺序是 `m:rPr` → `w:rPr` → `m:t`；`m:f` 是 `m:num` → `m:den`；
  `m:sSubSup` 是 `m:e` → `m:sub` → `m:sup`。顺序错了 Word 会报文档损坏。
* `<m:t>` 内容首尾有空格时要加 `xml:space="preserve"`（脚本已处理）。
* 公式字号跟随段落字号（模板正文 10.5pt），不要单独设 `w:sz`，否则和正文不协调。
