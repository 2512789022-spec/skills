# 版式参数（实测自最终稿 `学号姓名设计1实验报告.docx`）

所有数值由 python-docx 直接读出，可作为生成脚本的硬参数。改版时若参考稿不同，先跑 `scripts/inspect_reference.py` 重新取数。

## 1. 页面与默认样式

| 项 | 值 |
|---|---|
| 纸张 | A4，21.00 × 29.70 cm |
| 页边距 | 上 2.54、下 2.54、左 3.17、右 3.17 cm |
| 页眉距边 / 页脚距边 | 1.50 / 1.75 cm |
| 版心宽 | 14.66 cm |
| Normal 样式 | `sz=21`（10.5pt）、`szCs=22`、两端对齐、`widowControl=0`、`kern=2`；未设字体名（靠 run 级 rFonts） |

正文 run 统一带：`<w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>`、`<w:color w:val="000000"/>`，字号不写（继承 10.5pt）。

## 2. 封面（逐段）

| 元素 | 字号/字体 | 对齐 | 缩进 | 其他 |
|---|---|---|---|---|
| 报告名"电子线路综合设计实践实验报告" | 26pt，黑体 + Times New Roman，加粗 | 居中 | 0.64cm | 行距 1.5，段后 7.8pt |
| 校徽图片 | 2.77 × 2.75 cm，内嵌 | 居中 | — | 段后 7.8pt |
| "实验一：      语音放大电路的设计" | 18pt，加粗 | 左 | 2.54cm | 段后 7.8pt；"实验一："与题目之间用 6 个空格 |
| 姓　名 / 学　号 / 班　级 / 学　院 / 指导老师 | 16pt，宋体 + Times New Roman | 默认 | 2.47cm | 标签为"姓"+4 空格+"名："，值前用空格推到同一列；段后 7.8pt |
| 日期"二Ｏ二六年 10 月 21 日" | 18pt，加粗 | 居中 | — | 用全角"Ｏ" |

封面字段的横线（若模板用下划线空格）长度必须一致：改成 5 行 2 列无边框表格、右列只给底边加线，线宽=列宽。

## 3. 章节标题

一级（`1、`～`5、`）：

```
keepNext + 行距 1.5 + 左对齐 + outlineLvl=0
段前 8pt（第 1 章可 0）、段后 2pt
run：加粗、宋体 + Times New Roman、14pt、黑色
```

二级（`N.M `）：

```
keepNext + 行距 1.5 + 左对齐 + outlineLvl=1
段前 6pt、段后 2pt
run：加粗、10.5pt（不写字号，继承 Normal）
```

标题文本用全角顿号 `1、设计内容`，编号与文字之间不加空格；二级用半角 `2.1 各级电路的工作原理`（编号后一个空格）。

## 4. 正文与条目段

| 类型 | 对齐 | 行距 | 首行缩进 | 段前/段后 |
|---|---|---|---|---|
| 普通正文 | 两端对齐（默认） | 1.5（`line=360 auto`） | 0.74cm / `firstLineChars=200` `firstLine=420` | 0 |
| `（1）…（6）` 条目 | 两端对齐 | 1.5 | **0**（顶格） | 0 |
| 图题 | 居中 | 1.5 | 0 | 0 |
| 表题 | 居中 | 1.5 | 0 | 前 8pt、后 3pt |
| 公式段 | 居中 | 1.5 | 0 | 前 4pt、后 4pt |

图题/表题字号与正文相同（10.5pt），不加粗。题注格式：`图12 说明文字`、`表4 说明文字`——编号与说明之间一个半角空格。

## 5. 图片

- 单张仿真/示波器截图：宽 12.2cm（个别 12.06～12.79cm），图片段落居中，`space_before=2pt`（紧接上一段的图可 0）
- 章首实物全景照：10.87cm
- 校徽：2.77cm
- 子图标签 `(a)(b)`：最终稿用浮动文本框（1.17 × 3.90cm）叠在图上；**生成稿建议改用单元格首段/图前小段居中放标签**，跨软件更稳
- 多子图并排：2×2 无框表格（`tblStyle=Table Grid` 但边框可去掉），单元格内先标签后图片，表下一条图题

## 6. 三线表 XML（照抄即可）

```xml
<w:tblPr>
  <w:tblW w:w="8164" w:type="dxa"/>
  <w:jc w:val="center"/>
  <w:tblBorders>
    <w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/>
    <w:left w:val="none" w:sz="0" w:space="0" w:color="000000"/>
    <w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>
    <w:right w:val="none" w:sz="0" w:space="0" w:color="000000"/>
    <w:insideH w:val="none" w:sz="0" w:space="0" w:color="000000"/>
    <w:insideV w:val="none" w:sz="0" w:space="0" w:color="000000"/>
  </w:tblBorders>
  <w:tblLayout w:type="fixed"/>
  <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
</w:tblPr>
```

栏目行每个单元格再加：

```xml
<w:tcBorders><w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/></w:tcBorders>
```

单元格段落：居中、`spacing before=20 after=20`（twips，即 1pt）；栏目行 run 加粗。列宽用 `w:gridCol` 与 `tcW` 同时设，总和 ≤ 14.66cm。

## 7. 公式

共 7 处，全部为独立成段的 OMML（`m:oMathPara` + `m:oMath`，`jc=center`）：

1. 前置级增益 `A_u1 = 1 + R12/R1`
2. 高通截止频率 `f_1 = 1/(2π√(R3R4C1C2))`
3. `K_1 = 1 + R6/R5`、`Q_1 = 1/(3−K_1)`
4. 低通截止频率 `f_2 = 1/(2π√(R7R8C3C4))`
5. 两级串联幅频特性（分式 + 根号 + 绝对值）
6. 整机增益 `A_u = A_u1·A_u2·A_u3·A_u4`
7. 输出耦合电容截止频率 `f = 1/(2πR_LC8)`

正文里不出现 `f1≈106Hz` 这类字符公式，改写为"高通截止频率约106Hz"。

## 8. 页脚

页脚居中放 PAGE 域（内容控件 sdt，样式 `a5`），PDF 上显示为居中页码。首节与后续节分别设置，避免出现空页脚。
