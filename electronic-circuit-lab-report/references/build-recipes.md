# python-docx 代码骨架

依赖：`python-docx`、`lxml`、`latex2mathml`；PDF 导出靠 Word COM（PowerShell）。

## 初始化与字体

```python
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = docx.Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.top_margin = sec.bottom_margin = Cm(2.54)
sec.left_margin = sec.right_margin = Cm(3.17)
doc.styles['Normal'].font.size = Pt(10.5)
doc.styles['Normal'].paragraph_format.alignment = AL.JUSTIFY

def style_run(r, size=10.5, bold=False):
    r.font.size = Pt(size)
    r.font.bold = bold
    rPr = r._element.get_or_add_rPr()
    rf = rPr.find(qn('w:rFonts')) or OxmlElement('w:rFonts')
    if rf.getparent() is None:
        rPr.insert(0, rf)
    for a in ('w:ascii', 'w:hAnsi', 'w:cs'):
        rf.set(qn(a), 'Times New Roman')
    rf.set(qn('w:eastAsia'), '宋体')
    c = OxmlElement('w:color'); c.set(qn('w:val'), '000000'); rPr.append(c)
```

## 正文与条目段

```python
def P(text, indent=True):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing = 1.5
    pf.space_before = pf.space_after = Pt(0)
    if indent:                      # 普通正文 2 字符缩进
        pf.first_line_indent = Pt(21)
        p._p.get_or_add_pPr().find(qn('w:ind')).set(qn('w:firstLineChars'), '200')
    style_run(p.add_run(text))
    return p
# 以（1）（2）开头的条目：P(text, indent=False)
```

## 标题（带大纲级别，导航窗格可见）

`outlineLvl` 必须插在 `jc` 之后、`divId/cnfStyle/rPr/sectPr/pPrChange` 之前。

```python
SUCC = [qn('w:divId'), qn('w:cnfStyle'), qn('w:rPr'), qn('w:sectPr'), qn('w:pPrChange')]

def set_outline(p, lvl):
    pPr = p._p.get_or_add_pPr()
    el = OxmlElement('w:outlineLvl'); el.set(qn('w:val'), str(lvl))
    for ch in pPr:
        if ch.tag in SUCC:
            ch.addprevious(el); return
    pPr.append(el)

def H(text, lvl=1):
    p = doc.add_paragraph()
    p.paragraph_format.alignment = AL.LEFT
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.keep_with_next = True
    if lvl == 1:
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        style_run(p.add_run(text), size=14, bold=True)
        set_outline(p, 0)
    else:
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        style_run(p.add_run(text), bold=True)
        set_outline(p, 1)
    return p
```

## 图 + 图题

```python
def FIG(path, caption, width=12.2, before=2):
    p = doc.add_paragraph()
    p.paragraph_format.alignment = AL.CENTER
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.first_line_indent = Pt(0)
    p.add_run().add_picture(path, width=Cm(width))
    c = doc.add_paragraph()
    c.paragraph_format.alignment = AL.CENTER
    c.paragraph_format.line_spacing = 1.5
    c.paragraph_format.first_line_indent = Pt(0)
    style_run(c.add_run(caption))
```

## 三线表

```python
from docx.enum.table import WD_TABLE_ALIGNMENT

def TABLE(caption, rows, widths):
    cap = doc.add_paragraph()
    cap.paragraph_format.alignment = AL.CENTER
    cap.paragraph_format.space_before = Pt(8)
    cap.paragraph_format.space_after = Pt(3)
    style_run(cap.add_run(caption))

    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    tblPr = t._tbl.tblPr
    b = OxmlElement('w:tblBorders')
    for side, val, sz in (('top','single','12'), ('left','none','0'), ('bottom','single','12'),
                          ('right','none','0'), ('insideH','none','0'), ('insideV','none','0')):
        e = OxmlElement('w:' + side)
        e.set(qn('w:val'), val); e.set(qn('w:sz'), sz)
        e.set(qn('w:space'), '0'); e.set(qn('w:color'), '000000')
        b.append(e)
    tblPr.append(b)
    lay = OxmlElement('w:tblLayout'); lay.set(qn('w:type'), 'fixed'); tblPr.append(lay)
    for j, w in enumerate(widths):                 # 总和 ≤ 14.66cm
        t._tbl.tblGrid.findall(qn('w:gridCol'))[j].set(qn('w:w'), str(int(Cm(w).twips)))
        for row in t.rows:
            row.cells[j].width = Cm(w)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = t.cell(i, j)
            cell.width = Cm(widths[j])
            p = cell.paragraphs[0]
            p.alignment = AL.CENTER
            p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(1)
            style_run(p.add_run(str(val)), bold=(i == 0))
            if i == 0:                              # 栏目行下框线 0.75pt
                tb = OxmlElement('w:tcBorders')
                e = OxmlElement('w:bottom')
                e.set(qn('w:val'), 'single'); e.set(qn('w:sz'), '6')
                e.set(qn('w:space'), '0'); e.set(qn('w:color'), '000000')
                tb.append(e); cell._tc.get_or_add_tcPr().append(tb)
    return t
```

## Word 原生公式（OMML）

```python
import latex2mathml.converter as l2m
from lxml import etree

XSL = r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL"
XS = etree.XSLT(etree.parse(XSL))

def EQ(tex):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.alignment = AL.CENTER
    pf.line_spacing = 1.5
    pf.space_before = pf.space_after = Pt(4)
    root = XS(etree.fromstring(l2m.convert(tex, display='block'))).getroot()
    if etree.QName(root).localname == 'oMathPara':
        para = root
    else:
        para = OxmlElement('m:oMathPara')
        para.append(root)
    # 居中
    pr = para.find(qn('m:oMathParaPr'))
    if pr is None:
        pr = OxmlElement('m:oMathParaPr'); para.insert(0, pr)
    if pr.find(qn('m:jc')) is None:
        jc = OxmlElement('m:jc'); pr.append(jc)
    pr.find(qn('m:jc')).set(qn('m:val'), 'center')
    p._p.append(para)
```

## 页脚页码

```python
p = sec.footer.paragraphs[0]
p.alignment = AL.CENTER
fld = OxmlElement('w:fldSimple'); fld.set(qn('w:instr'), ' PAGE ')
r = OxmlElement('w:r'); t = OxmlElement('w:t'); t.text = '1'
r.append(t); fld.append(r); p._p.append(fld)
```

## 保存与占用处理

```python
import os, shutil
def save(doc, target):
    tmp = '_build_tmp.docx'
    try:
        doc.save(tmp)
        os.replace(tmp, target)          # 目标未被占用时直接覆盖
        return target
    except PermissionError:
        alt = target.replace('.docx', '_new.docx')
        shutil.move(tmp, alt)            # 被 Word/预览占用 → 另存并告知用户
        return alt
```

## 校验与自查

```bash
python scripts/verify_report.py 报告.docx
```

导出 PDF 后建议用 pymupdf 渲染每页成 PNG，逐页目视；分页异常的典型症状是"上一页底部大段空白 + 整块图表跑到下一页"，改图宽或调段落间距即可解决。
