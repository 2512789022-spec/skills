---
name: electronic-circuit-lab-report
description: 生成、排版与修订《电子线路综合设计实践》等中文电子线路/模拟电路实验报告的 Word 文档（.docx）并导出 PDF。当用户要写/生成/美化/统一实验报告格式，涉及字体字号、1.5 倍行距、首行缩进、图号与图题、三线表、Word 原生公式、封面与页码时使用。This skill should be used when producing or restyling a Chinese electronics-lab report in .docx with fixed typography, three-line tables, numbered figures/captions and Word-native equations.
agent_created: true
---

# 电子线路实验报告（docx）生成与排版

## 用途

按固定版式产出一份可直接提交的中文实验报告：封面 → 大纲级别标题（导航窗格可见）→ 正文五章 → 图/表/公式 → 页码，然后导出同名 PDF。也用于把已有报告改版到这套规范，或按用户既有报告"对齐格式"。

## 交付契约

- 文件命名：`<学号><姓名>_<课程名>_<设计编号>_<题目>.docx`，同名 `.pdf` 放同一目录
- 生成稿另起文件名；**修订既有报告时另存新文件，绝不覆盖原件**
- 交付时同时给出 docx 与 pdf，说明页数与本次改动点

## 章节骨架（固定五章，编号沿用全角顿号）

```
1、设计内容            任务与指标 + 总体原理图
2、计算分析            2.1 各级电路的工作原理
                       2.2 关键性能参数计算
                       2.3 元件参数变化对性能的影响
3、仿真验证            3.1～3.4  与课程要求的仿真条目一一对应
4、焊接调试            4.1 起连续编号，只保留实际做过的实测项
5、总结
```

课程要求条目未实测时：**不预留占位、不编造数据**，删掉该小节并把交叉引用（如"结果见4.1～4.2"）同步改掉；仿真能做的 3.x 照旧保留。

## 格式规范速查

| 项目 | 规定 |
|---|---|
| 页面 | A4；上下 2.54cm，左右 3.17cm（版心 14.66cm） |
| 正文 | 中文宋体 / 西文 Times New Roman，小五 10.5pt，两端对齐，1.5 倍行距，首行缩进 0.74cm（=2 字符，`firstLineChars=200`） |
| 条目段 | 以 `（1）（2）…` 开头的段落首行**不缩进**（顶格） |
| 一级标题 | `N、标题`，加粗 14pt，左对齐，段前 8pt、段后 2pt，1.5 行距，`outlineLvl=0` |
| 二级标题 | `N.M 标题`，加粗 10.5pt，左对齐，段前 6pt、段后 2pt，`outlineLvl=1` |
| 图 | 图片段落居中、宽 **12.2cm**；图题居中 `图N 说明`，1.5 行距、无缩进、置于图下 |
| 表 | 三线表：上下框线 1.5pt、栏目行下框线 0.75pt、无竖线；表宽固定、整体居中；表题 `表N 说明` 居中、段前 8pt 段后 3pt、置于表上 |
| 公式 | Word 原生 OMML，独立成段居中；**不得用字符拼公式** |
| 页码 | 页脚居中 PAGE 域 |
| 封面 | 26pt 黑体标题 + 校徽 + 18pt 加粗"实验N：题目" + 5 个 16pt 字段 + 18pt 日期；字段行缩进 2.47cm、用全角空格对齐；详见 `references/format-spec.md` |

## 实施流程

1. **取指纹**：有参考稿时先跑 `python scripts/inspect_reference.py <参考.docx>`，以实测数值为准，不要凭记忆写格式。
2. **生成**：用 python-docx 脚本一次成稿，骨架见 `references/build-recipes.md`。
3. **自检**：`python scripts/verify_report.py <成品.docx>`，全部 PASS 再往下走。
4. **导出 PDF**：用 PowerShell 工具调 Word COM（见下），记下页数。
5. **目视抽查**：把 PDF 每页渲染成图，重点看封面、每章首页、含图表页、末页；出现孤行、大块空白、图题错位就回改。
6. **交付**：`present_files` 同时给 docx 与 pdf，正文说明页数、章节结构、本次改动。

### 导出 PDF（Windows，Word COM）

```powershell
$doc="C:\path\报告.docx"; $pdf="C:\path\报告.pdf"
if (Test-Path $pdf) { Remove-Item $pdf -Force }
$w = New-Object -ComObject Word.Application
$w.Visible = $false; $w.DisplayAlerts = 0
$d = $w.Documents.Open($doc, $false, $true)
$d.ExportAsFixedFormat($pdf, 17)
Write-Output ("pages=" + $d.ComputeStatistics(2))
$d.Close(0); $w.Quit()
```

## 关键坑（逐条核对）

1. **文件被占用**：Word/预览打开时会 `PermissionError` / `WinError 5`。先写临时副本再 `os.replace`；仍失败就另存新文件名，并明确告知用户"原文件被占用"。
2. **deepcopy 段落**会带重复 `w14:paraId` / `w14:textId`，复制后删掉这两个属性，否则 Word 可能提示修复。
3. **表格跨页断开**（标签留在页底、图片跑到次页）：行加 `<w:cantSplit/>`，并把前一行所有段落、本行首段设 `keep_with_next`，整表锁在一页。
4. **列宽**：总和不大于 14.66cm，且设 `tblLayout=fixed`，否则自动布局会撑出版心。
5. **图片宽度**改 0.4cm 就可能引起整章翻页：统一 12.2cm，改完必须重新导出 PDF 看分页。
6. **公式**必须转 OMML：`latex2mathml` 生成 MathML，再用 Office 的 `MML2OMML.XSL`（`C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL`）做 XSLT；`oMathPara` 设 `m:jc=center`，且不要把 `oMathPara` 套进 `oMath`。
7. **图号/表号**连续、正文引用同步；删图后全文搜 `图\d` 逐处核对，不留悬空引用。
8. **子图标注**：同一组多张子图用 `(a)(b)(c)(d)`，标签在图片上方（单元格首段或图前小段），整组只给**一条**图题。
9. **修订既有报告**：只改指定段落；改前改后按段落文本 diff 确认其余零改动；被改段的 `rPr/pPr` 用 deepcopy 原段再换文字，保证逐字节一致。
10. **中文排版**：中西文之间不留字面空格（靠 Word autoSpace）；数值与单位紧贴（`5mVpp`、`1kHz`、`1000μF`）；范围用 `～`，因果用 `→`，算式等号用全角 `＝`。

## 资源

- `references/format-spec.md` — 从既有最终稿实测出的精确参数（含 XML 片段、封面尺寸）
- `references/content-guide.md` — 各章内容要求、数值一致性、图表与引用规范、扣分点
- `references/build-recipes.md` — python-docx 代码片段：字体/标题/正文/图/三线表/公式/页码/封面
- `scripts/inspect_reference.py` — 提取任意 docx 的版式指纹
- `scripts/verify_report.py` — 成品格式自检（字号、缩进、行距、图号连续性、三线表、公式）
