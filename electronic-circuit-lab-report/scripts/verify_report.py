#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""检查成品实验报告是否符合既定版式；不合规项逐条列出，有 FAIL 则退出码 1。

用法:
    python verify_report.py <成品.docx> [--body-start 索引]

未给 --body-start 时，自动以第一个匹配 `N、` 的标题为正文起点（封面不参与字体/缩进检查）。
"""
import argparse
import re
import sys

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
from docx.oxml.ns import qn

PASS, FAIL = 'PASS', 'FAIL'
results = []


def check(ok, msg, detail=''):
    results.append((PASS if ok else FAIL, msg, detail))


def cm(v):
    return None if v is None else round(v.cm, 2)


def run_props(p):
    out = []
    for r in p.runs:
        if not r.text:
            continue
        rPr = r._element.rPr
        rf = rPr.find(qn('w:rFonts')) if rPr is not None else None
        out.append({
            'text': r.text[:12],
            'size': r.font.size.pt if r.font.size else None,
            'bold': r.font.bold,
            'ascii': rf.get(qn('w:ascii')) if rf is not None else None,
            'east': rf.get(qn('w:eastAsia')) if rf is not None else None,
        })
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('path')
    ap.add_argument('--body-start', type=int, default=None)
    a = ap.parse_args()
    d = docx.Document(a.path)

    # ---------- 页面 ----------
    s = d.sections[0]
    check(abs(s.page_width.cm - 21) < 0.05 and abs(s.page_height.cm - 29.7) < 0.05,
          'A4 纸张 21x29.7cm', f'{cm(s.page_width)}x{cm(s.page_height)}')
    check(all(abs(x.cm - y) < 0.02 for x, y in
              ((s.top_margin, 2.54), (s.bottom_margin, 2.54), (s.left_margin, 3.17), (s.right_margin, 3.17))),
          '页边距 2.54/2.54/3.17/3.17cm',
          f'{cm(s.top_margin)}/{cm(s.bottom_margin)}/{cm(s.left_margin)}/{cm(s.right_margin)}')
    st = d.styles['Normal']
    check(st.font.size is not None and abs(st.font.size.pt - 10.5) < 0.01,
          'Normal 字号 10.5pt（小五）', str(st.font.size.pt if st.font.size else None))

    # ---------- 正文起点 ----------
    start = a.body_start
    if start is None:
        start = next((i for i, p in enumerate(d.paragraphs) if re.match(r'^[1-9]、\S', p.text.strip())), 0)
    print(f'正文起点: 段 {start}\n')

    h1 = re.compile(r'^[1-9]、\S')
    h2 = re.compile(r'^\d\.\d+\s')
    cap_fig = re.compile(r'^图\s?\d+\s')
    cap_tab = re.compile(r'^表\s?\d+\s')
    entry = re.compile(r'^（\d+）')

    bad_h1, bad_h2, bad_body, bad_entry, bad_cap, bad_font = [], [], [], [], [], []
    figs, tabs = [], []

    for i, p in enumerate(d.paragraphs):
        if i < start:
            continue
        t = p.text.strip()
        if not t:
            continue
        pf = p.paragraph_format
        pPr = p._p.pPr
        lvl = pPr.find(qn('w:outlineLvl')) if pPr is not None else None
        lvl = lvl.get(qn('w:val')) if lvl is not None else None
        rps = run_props(p)

        if i >= start:
            for rp in rps:
                if rp['east'] not in ('宋体', '黑体'):
                    bad_font.append((i, rp['text'], rp['east']))
                elif rp['ascii'] not in ('Times New Roman',) and rp['ascii'] is not None:
                    bad_font.append((i, rp['text'], rp['ascii']))
                if rp['size'] not in (None, 10.5, 14.0):
                    bad_font.append((i, rp['text'], f"{rp['size']}pt"))

        if h1.match(t):
            ok = (lvl == '0' and any(r['bold'] for r in rps)
                  and all(r['size'] in (None, 14.0) for r in rps)
                  and pf.alignment in (None, AL.LEFT))
            if not ok:
                bad_h1.append((i, t[:20], lvl, [r['size'] for r in rps]))
        elif h2.match(t):
            ok = (lvl == '1' and any(r['bold'] for r in rps)
                  and pf.alignment in (None, AL.LEFT))
            if not ok:
                bad_h2.append((i, t[:20], lvl))
        elif cap_fig.match(t) or cap_tab.match(t):
            (figs if cap_fig.match(t) else tabs).append((i, t))
            ok = (pf.alignment == AL.CENTER and pf.first_line_indent in (None, 0))
            if not ok:
                bad_cap.append((i, t[:20], str(pf.alignment), cm(pf.first_line_indent)))
        elif entry.match(t):
            ind = pf.first_line_indent
            if ind is not None and abs(ind.cm) > 0.01:
                bad_entry.append((i, t[:20], cm(ind)))
        else:
            ind = pf.first_line_indent
            want = 0.74
            if ind is None or abs(ind.cm - want) > 0.03:
                bad_body.append((i, t[:20], cm(ind)))
            if pf.line_spacing not in (1.5, None):
                bad_body.append((i, t[:20], f'line={pf.line_spacing}'))

    check(not bad_h1, '一级标题 `N、` 加粗14pt/左对齐/大纲1级', str(bad_h1[:3]))
    check(not bad_h2, '二级标题 `N.M ` 加粗/左对齐/大纲2级', str(bad_h2[:3]))
    check(not bad_body, '正文 2 字符缩进（0.74cm）+ 1.5 倍行距', str(bad_body[:3]))
    check(not bad_entry, '`（N）` 条目段顶格（无首行缩进）', str(bad_entry[:3]))
    check(not bad_cap, '图表题注居中且不缩进', str(bad_cap[:3]))
    check(not bad_font, '字体：中文宋体+西文Times New Roman，字号仅 10.5/14pt', str(bad_font[:3]))

    # ---------- 图号连续 ----------
    if figs:
        nums = [int(re.match(r'^图\s?(\d+)', t).group(1)) for _, t in figs]
        check(nums == list(range(1, len(nums) + 1)), '图号连续（1..N）', str(nums))
        text = '\n'.join(p.text for p in d.paragraphs)
        for t in d.tables:
            for row in t.rows:
                for c in row.cells:
                    text += '\n' + c.text
        refs = {int(m) for m in re.findall(r'图\s?(\d+)', text)}
        missing = sorted(refs - set(nums))
        check(not missing, '正文引用的图号都存在（无悬空引用）', str(missing))
    else:
        check(False, '未找到 `图N 说明` 形式的图题', '')

    if tabs:
        tn = [int(re.match(r'^表\s?(\d+)', t).group(1)) for _, t in tabs]
        check(tn == list(range(1, len(tn) + 1)), '表号连续（1..N）', str(tn))

    # ---------- 图片宽度 ----------
    widths, odd = [], []
    for i, p in enumerate(d.paragraphs):
        if i < start:
            continue
        for e in p._p.findall('.//' + qn('wp:extent')):
            w = int(e.get('cx')) / 360000
            widths.append(round(w, 2))
            if w > 3 and not (10.5 <= w <= 13.2):
                odd.append((i, round(w, 2)))
    for ti, t in enumerate(d.tables):
        for e in t._tbl.findall('.//' + qn('wp:extent')):
            w = int(e.get('cx')) / 360000
            widths.append(round(w, 2))
    check(not odd, '插图宽度 10.5～13.2cm（正文图 12.2cm）', str(odd[:5]))
    print('  图片宽度:', widths)

    # ---------- 三线表 ----------
    bad_tbl = []
    for ti, t in enumerate(d.tables):
        if len(t._tbl.findall('.//' + qn('a:blip'))) == 4:   # 照片排布表不算
            continue
        tb = t._tbl.tblPr.find(qn('w:tblBorders'))
        if tb is None:
            bad_tbl.append((ti, 'no borders'))
            continue
        got = {c.tag.split('}')[1]: (c.get(qn('w:val')), c.get(qn('w:sz'))) for c in tb}
        ok = (got.get('top') == ('single', '12') and got.get('bottom') == ('single', '12')
              and all(got.get(k, ('none',))[0] == 'none' for k in ('left', 'right', 'insideH', 'insideV')))
        if not ok:
            bad_tbl.append((ti, got))
    check(not bad_tbl, '表格为三线表（上下 1.5pt、无竖线）', str(bad_tbl[:3]))

    # ---------- 公式 ----------
    n_math = len(d.element.body.findall('.//' + qn('m:oMath')))
    print(f'  公式（OMML）数量: {n_math}')

    # ---------- 汇总 ----------
    print('\n== 结果 ==')
    for st_, msg, detail in results:
        line = f'[{st_}] {msg}'
        if st_ == FAIL and detail:
            line += f'  →  {detail}'
        print(line)
    fails = sum(1 for r in results if r[0] == FAIL)
    print(f'\n共 {len(results)} 项，FAIL {fails} 项')
    return 1 if fails else 0


if __name__ == '__main__':
    sys.exit(main())
