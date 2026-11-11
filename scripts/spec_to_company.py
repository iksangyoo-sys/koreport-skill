#!/usr/bin/env python3
"""
koreport-skill — report_spec.md를 회사 양식(templates/row.pptx 가로, col.pptx A4 세로)으로 바로 렌더링한다.

AI 없이 PC에서 실행하므로 토큰이 들지 않는다. 문법은 spec_to_pptx.py와 같은
report_spec.md(workflows/02-spec-format.md)를 그대로 쓴다.

사용법:
    python spec_to_company.py <report_spec.md> [-o 결과.pptx] [--layout row|col] [--check]

    --layout row   가로 양식(기본). '## 장 제목' 하나 = 장표 한 장, 표지는 양식 1쪽을 채운다.
    --layout col   A4 세로 양식. 제목 배너 아래로 '## 장'을 '1. 장 제목 : 헤드라인' 섹션으로
                   차례로 쌓고, 쪽이 차면 다음 쪽으로 넘긴다. '## [새쪽] …'은 새 쪽에서 시작.
    머리말에 'layout: col'을 적어도 된다(명령행 옵션이 우선).

강조: '**핵심어**' → 가로는 #3333CC 굵게, 세로는 #1F77A2 굵게.
세로 양식에서 【라벨】 바로 아래 줄들은 '회색 둥근 라벨 + 본문' 행으로 그린다.
의존성: python-pptx
"""

from __future__ import annotations

import argparse
import copy
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from spec_to_pptx import Block, Slide, Spec, check, parse_spec  # noqa: E402

from lxml import etree  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.dml.color import RGBColor  # noqa: E402
from pptx.enum.shapes import MSO_SHAPE  # noqa: E402
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402
from pptx.util import Inches, Pt  # noqa: E402

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

TEMPLATES = HERE.parent / "templates"
FONT = "맑은 고딕"
BLACK = RGBColor(0, 0, 0)
NAVY = RGBColor(0x00, 0x20, 0x60)
GRAY = RGBColor(0x59, 0x59, 0x59)
LINE = RGBColor(0xA6, 0xA6, 0xA6)
LGRAY = RGBColor(0xD9, 0xD9, 0xD9)

LEVEL1 = ("❍", "■", "○", "□", "◎")


def I(v: float) -> int:
    return int(Inches(v))


# ── 공통 텍스트 유틸 ───────────────────────────────────────────────────
def runs(p, text: str, size: float, color=BLACK, bold=False, key=NAVY):
    for part in re.split(r"(\*\*[^*]+\*\*)", text):
        if not part:
            continue
        c, b = color, bold
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            part, c, b = part[2:-2], key, True
        r = p.add_run()
        r.text = part
        f = r.font
        f.size, f.bold, f.name = Pt(size), b, FONT
        f.color.rgb = c
        rpr = r._r.get_or_add_rPr()
        for tag in ("a:latin", "a:ea"):
            el = rpr.find(qn(tag))
            if el is None:
                el = rpr.makeelement(qn(tag), {})
                rpr.append(el)
            el.set("typeface", FONT)


def para_fmt(p, line=1.1, before=0.0, align=PP_ALIGN.LEFT, hang=0.0, left=0.0):
    p.line_spacing = line
    p.space_before = Pt(before)
    p.alignment = align
    if hang or left:
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", str(I(left + hang)))
        pPr.set("indent", str(-I(hang)))


def tbox(s, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    tb = s.shapes.add_textbox(I(x), I(y), I(w), max(I(h), I(0.15)))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = I(0.05)
    tf.margin_top = tf.margin_bottom = I(0.02)
    return tf


def _drop_style(shp):
    st = shp._element.find(qn("p:style"))
    if st is not None:
        shp._element.remove(st)


def outline_box(s, x, y, w, h):
    b = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, I(x), I(y), I(w), I(h))
    _drop_style(b)
    b.fill.background()
    b.line.color.rgb = LINE
    b.line.width = Pt(0.25)
    return b


def hline(s, x, y, w, color=LINE, width=0.5):
    ln = s.shapes.add_connector(1, I(x), I(y), I(x + w), I(y))
    _drop_style(ln)
    ln.line.color.rgb = color
    ln.line.width = Pt(width)
    return ln


def text_w(text: str, size: float) -> float:
    """글자 폭 합(인치). 한글·전각 ≈ 1em, 영문·숫자 ≈ 0.55em."""
    t = re.sub(r"\*\*", "", text)
    em = sum(1.0 if ord(ch) > 0x2E80 else 0.55 for ch in t)
    return em * size / 72


def text_h(text: str, size: float, width: float, line=1.1) -> float:
    n = max(1, -(-text_w(text, size) // max(width - 0.15, 0.5)))
    return n * size * 1.2 * line / 72


def set_para_text(p, text: str):
    rs = p.runs
    if rs:
        rs[0].text = text
        for r in rs[1:]:
            p._p.remove(r._r)
    else:
        p.add_run().text = text


def table(s, rows, x, y, w, size, head_hex="D9D9D9", head_color=NAVY, key=NAVY):
    ncol = max(len(r) for r in rows)
    rows = [r + [""] * (ncol - len(r)) for r in rows]
    row_h = size * 2.0 / 72
    gt = s.shapes.add_table(len(rows), ncol, I(x), I(y), I(w), I(row_h * len(rows)))
    tbl = gt.table
    weights = [max(0.6, max(text_w(r[c], size) for r in rows)) for c in range(ncol)]
    for c, wt in enumerate(weights):
        tbl.columns[c].width = int(I(w) * wt / sum(weights))
    for ri in range(len(rows)):
        tbl.rows[ri].height = I(row_h)
    tblpr = tbl._tbl.tblPr
    tblpr.set("firstRow", "0")
    tblpr.set("bandRow", "0")
    num = re.compile(r"^[\s△▲▼↑↓+\-−]?[\d,.]+\s*(%p?|억|천|만|원|명|건|개|시간|배|社)?[\s↑↓]*$")
    for ri, r in enumerate(rows):
        for ci, val in enumerate(r):
            cell = tbl.cell(ri, ci)
            cell.margin_left = cell.margin_right = I(0.05)
            cell.margin_top = cell.margin_bottom = I(0.01)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            if ri == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor.from_string(head_hex)
                runs(p, val, size, head_color, True, key=head_color)
                p.alignment = PP_ALIGN.CENTER
            else:
                cell.fill.background()
                runs(p, val, size, key=key)
                p.alignment = PP_ALIGN.CENTER if num.match(val or "x") else PP_ALIGN.LEFT
            tcpr = cell._tc.get_or_add_tcPr()
            for pos, side in enumerate(("L", "R", "T", "B")):
                tag = qn(f"a:ln{side}")
                old = tcpr.find(tag)
                if old is not None:
                    tcpr.remove(old)
                on = side in ("T", "B")
                ln = tcpr.makeelement(tag, {"w": "3175" if on else "0"})
                if on:
                    sf = ln.makeelement(qn("a:solidFill"), {})
                    sf.append(sf.makeelement(qn("a:srgbClr"), {"val": "A6A6A6"}))
                    ln.append(sf)
                else:
                    ln.append(ln.makeelement(qn("a:noFill"), {}))
                tcpr.insert(pos, ln)
    return row_h * len(rows)


def drop_slides(prs, keep: set[int]):
    lst = prs.slides._sldIdLst
    for i, sld in reversed(list(enumerate(list(lst)))):
        if i not in keep:
            prs.part.drop_rel(sld.get(qn("r:id")))
            lst.remove(sld)


# ── 가로 양식 (row.pptx) ──────────────────────────────────────────────
ROW_KEY = RGBColor(0x33, 0x33, 0xCC)
ROW_W, ROW_H = 10.83, 7.5
ROW_L, ROW_CW = 0.37, 10.1


def row_style(t: str):
    """기호별 (크기, 굵게, 색, 왼쪽 들여쓰기, 매달림, 앞 간격pt)."""
    if t.startswith(LEVEL1) or re.match(r"^[①-⑩❶-❿]", t):
        return 15, True, BLACK, 0.0, 0.25, 9
    if t.startswith("•"):
        return 13, False, BLACK, 0.2, 0.17, 3
    if t.startswith(("-", "–")):
        return 12, False, GRAY, 0.4, 0.15, 1
    if t.startswith("☞"):
        return 14, True, BLACK, 0.0, 0.25, 7
    if t.startswith("※"):
        return 10.5, False, GRAY, 0.0, 0.2, 3
    if t.startswith("▷"):
        return 12.5, False, NAVY, 0.2, 0.2, 3
    return 13, False, BLACK, 0.0, 0.0, 3


GAPF = {"v": 1.0}   # 내용이 적은 장은 줄 간격을 넓혀 빈 공간을 줄인다


def row_lines_h(lines, w, scale):
    h = 0.0
    for t in lines:
        size, _, _, left, _, before = row_style(t)
        h += text_h(t, size * scale, w - left) + before * scale * GAPF["v"] / 72
    return h + 0.06


def row_block_h(blk: Block, w, scale):
    h = 0.0
    if blk.title:
        h += 0.38 + 0.08
    if blk.table:
        h += 11.5 * scale * 2.0 / 72 * len(blk.table) + 0.08
    if blk.lines:
        h += row_lines_h(blk.lines, w - (0.24 if blk.title else 0), scale)
    if blk.title:
        h += 0.08
    return h + 0.1


def row_draw_lines(s, x, y, w, lines, scale):
    h = row_lines_h(lines, w, scale)
    tf = tbox(s, x, y, w, h)
    for i, t in enumerate(lines):
        size, bold, color, left, hang, before = row_style(t)
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para_fmt(p, 1.1, 0 if i == 0 else before * scale * GAPF["v"], hang=hang, left=left)
        runs(p, t, size * scale, color, bold, key=ROW_KEY)
    return h


def row_draw_block(s, blk: Block, x, y, w, scale):
    y0 = y
    inner_x, inner_w = x, w
    if blk.title:
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, I(x), I(y), I(w), I(0.38))
        _drop_style(bar)
        bar.fill.solid()
        bar.fill.fore_color.rgb = RGBColor(0xF2, 0xF2, 0xF2)
        bar.line.fill.background()
        tf = bar.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        para_fmt(p, 1.0, align=PP_ALIGN.CENTER)
        runs(p, blk.title.strip("【】 "), 12.5 * scale, NAVY, True, key=NAVY)
        y += 0.38 + 0.08
        inner_x, inner_w = x + 0.12, w - 0.24
    if blk.table:
        y += table(s, blk.table, inner_x, y, inner_w, 11.5 * scale, key=ROW_KEY) + 0.08
    if blk.lines:
        y += row_draw_lines(s, inner_x, y, inner_w, blk.lines, scale)
    if blk.title:
        y += 0.08
        outline_box(s, x, y0, w, y - y0)
    return y - y0 + 0.1


def row_body(prs, layout, sl: Slide, page: str):
    s = prs.slides.add_slide(layout)
    for ph in list(s.placeholders):
        if ph.placeholder_format.idx != 0:
            ph._element.getparent().remove(ph._element)
    title = re.sub(r"^\[(첨부|참고)(\d*)\]\s*", lambda m: f"[{m.group(1)}{m.group(2)}] ", sl.chapter)
    title = re.sub(r"\s+_\s+", " : ", title)
    s.shapes.title.text_frame.text = title
    if page:
        tf = tbox(s, 9.96, 0.31, 0.79, 0.4, MSO_ANCHOR.MIDDLE)
        p = tf.paragraphs[0]
        para_fmt(p, 1.0, align=PP_ALIGN.CENTER)
        cur, tot = page.split("/")
        runs(p, f"**{cur}**/{tot}", 13, GRAY, key=GRAY)

    if sl.kind == "section":
        tf = tbox(s, 0.8, 3.0, ROW_W - 1.6, 1.2, MSO_ANCHOR.MIDDLE)
        p = tf.paragraphs[0]
        para_fmt(p, 1.0, align=PP_ALIGN.CENTER)
        runs(p, sl.chapter, 28, NAVY, True)
        return s

    top = 0.9
    if sl.headline:
        h = text_h(sl.headline, 16, ROW_CW) + 0.08
        tf = tbox(s, ROW_L, top, ROW_CW, h, MSO_ANCHOR.MIDDLE)
        p = tf.paragraphs[0]
        para_fmt(p, 1.05, align=PP_ALIGN.CENTER)
        runs(p, sl.headline, 16, BLACK, True, key=ROW_KEY)
        top += h + 0.04
    if sl.subline:
        h = text_h(sl.subline, 11, ROW_CW) + 0.04
        tf = tbox(s, ROW_L, top, ROW_CW, h)
        p = tf.paragraphs[0]
        para_fmt(p, 1.0, align=PP_ALIGN.CENTER)
        runs(p, sl.subline, 11, GRAY, key=ROW_KEY)
        top += h + 0.04
    if sl.headline or sl.subline:
        hline(s, ROW_L, top + 0.02, ROW_CW)
        top += 0.12

    bottom_h = 0.45 if sl.bottom else 0
    foot_h = 0.24 * len(sl.footnotes)
    body_bottom = ROW_H - 0.25 - bottom_h - foot_h
    ncol = len(sl.columns)
    gap = 0.25
    col_w = (ROW_CW - gap * (ncol - 1)) / ncol

    # 넘치면 글자 크기를 최대 25%까지 줄인다.
    GAPF["v"] = 1.0
    need = max(sum(row_block_h(b, col_w, 1.0) for b in col) for col in sl.columns) if sl.columns else 0
    avail = body_bottom - top
    if 0 < need < avail * 0.55:
        GAPF["v"] = 2.0
    scale = 1.0 if need <= avail else max(0.75, avail / need)
    if need * scale > avail + 0.05:
        print(f"  ! '{sl.chapter}' 장: 내용이 많아 넘칠 수 있음 — 줄이거나 장을 나눌 것")

    for ci, col in enumerate(sl.columns):
        x = ROW_L + ci * (col_w + gap)
        y = top
        for blk in col:
            y += row_draw_block(s, blk, x, y, col_w, scale)
        if ncol > 1 and ci < ncol - 1:
            lx = x + col_w + gap / 2
            ln = s.shapes.add_connector(1, I(lx), I(top), I(lx), I(body_bottom))
            _drop_style(ln)
            ln.line.color.rgb = LINE
            ln.line.width = Pt(0.5)
            lnel = ln.line._get_or_add_ln()
            etree.SubElement(lnel, qn("a:prstDash"), val="sysDot")

    fy = body_bottom + 0.03
    for fn in sl.footnotes:
        tf = tbox(s, ROW_L, fy, ROW_CW, 0.24)
        runs(tf.paragraphs[0], fn, 10, GRAY, key=GRAY)
        fy += 0.24
    if sl.bottom:
        tf = tbox(s, 0.0, ROW_H - 0.25 - bottom_h + 0.02, ROW_W, bottom_h, MSO_ANCHOR.MIDDLE)
        p = tf.paragraphs[0]
        para_fmt(p, 1.0, align=PP_ALIGN.CENTER)
        runs(p, f"“{sl.bottom}”", 17, NAVY, True, key=NAVY)
    if sl.notes:
        s.notes_slide.notes_text_frame.text = "\n".join(sl.notes)
    return s


def row_cover(s, meta: dict, toc: list[str]):
    shapes = {sh.name: sh for sh in s.shapes}
    if "Text Box 9" in shapes:
        tb = shapes["Text Box 9"]
        p = tb.text_frame.paragraphs[0]
        title = meta.get("title", "")
        set_para_text(p, title)
        r = p.runs[0]
        size = (r.font.size.pt if r.font.size else 32)
        fit = (tb.width / 914400 - 0.3) * 72 / max(text_w(title, 1) * 72, 1)
        if fit < size:
            r.font.size = Pt(max(18, int(fit)))
    if "Rectangle 3" in shapes:
        ps = shapes["Rectangle 3"].text_frame.paragraphs
        set_para_text(ps[0], meta.get("date", ""))
        if len(ps) > 1:
            set_para_text(ps[1], meta.get("dept", ""))
    if "Rectangle 13" in shapes and toc:
        tf = shapes["Rectangle 13"].text_frame
        ps = tf.paragraphs
        for i, item in enumerate(toc):
            if i < len(ps):
                set_para_text(ps[i], item)
            else:
                new = copy.deepcopy(ps[-1]._p)
                tf._txBody.append(new)
                set_para_text(tf.paragraphs[-1], item)
        for p in list(tf.paragraphs)[len(toc):]:
            tf._txBody.remove(p._p)
    if meta.get("subtitle") and "Text Box 9" in shapes:
        tb = shapes["Text Box 9"]
        tf = tbox(s, tb.left / 914400, (tb.top + tb.height) / 914400, tb.width / 914400, 0.4)
        p = tf.paragraphs[0]
        para_fmt(p, 1.0, align=PP_ALIGN.CENTER)
        runs(p, meta["subtitle"], 16, NAVY)


def render_row(spec: Spec, template: Path, out: Path):
    prs = Presentation(template)
    layout = prs.slides[2].slide_layout   # 본문 장(그라데이션 제목 바)
    cover = spec.meta.get("cover", "yes").lower() not in ("no", "false", "0")
    drop_slides(prs, {0} if cover else set())
    if cover:
        toc = [t.strip() for t in spec.meta.get("toc", "").split("|") if t.strip()]
        row_cover(prs.slides[0], spec.meta, toc)
    body = [sl for sl in spec.slides if sl.kind == "body"]
    k = 0
    for sl in spec.slides:
        if sl.kind == "body":
            k += 1
            row_body(prs, layout, sl, f"{k}/{len(body)}")
        else:
            row_body(prs, layout, sl, "")
    prs.save(out)


# ── 세로 양식 (col.pptx, A4) ──────────────────────────────────────────
COL_KEY = RGBColor(0x1F, 0x77, 0xA2)
COL_W, COL_H = 7.5, 10.0
COL_X, COL_BW = 0.29, 6.93          # 섹션 박스
COL_TOP1, COL_TOPN, COL_BOTTOM = 0.82, 0.45, 9.65


def col_style(t: str):
    if t.startswith(LEVEL1) or re.match(r"^[①-⑩❶-❿]", t):
        return 12, True, BLACK, 0.0, 0.18, 4
    if t.startswith("•"):
        return 11.5, False, BLACK, 0.12, 0.15, 2
    if t.startswith(("-", "–")):
        return 10.5, False, GRAY, 0.3, 0.13, 1
    if t.startswith("☞"):
        return 12, True, COL_KEY, 0.0, 0.2, 4
    if t.startswith("※"):
        return 10, False, GRAY, 0.0, 0.17, 2
    if t.startswith("▷"):
        return 11, False, BLACK, 0.12, 0.17, 2
    return 11.5, False, BLACK, 0.0, 0.0, 2


def col_lines_h(lines, w):
    return sum(text_h(t, col_style(t)[0], w - col_style(t)[3], 1.05) + col_style(t)[5] / 72 for t in lines) + 0.04


def col_draw_lines(s, x, y, w, lines):
    h = col_lines_h(lines, w)
    tf = tbox(s, x, y, w, h)
    for i, t in enumerate(lines):
        size, bold, color, left, hang, before = col_style(t)
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para_fmt(p, 1.05, 0 if i == 0 else before, hang=hang, left=left)
        runs(p, t, size, color, bold, key=COL_KEY)
    return h


PILL_W = 0.98


def col_block_h(blk: Block) -> float:
    h = 0.0
    if blk.table:
        h += 10.5 * 2.0 / 72 * len(blk.table) + 0.08
    if blk.lines:
        w = COL_BW - 0.22 - (PILL_W + 0.1 if blk.title else 0)
        h += max(col_lines_h(blk.lines, w), 0.28 if blk.title else 0)
    elif blk.title and not blk.table:
        h += 0.28
    if blk.title and blk.table:
        h += 0.3
    return h + 0.08


def col_section_h(sl: Slide) -> float:
    head = sl.chapter + (" : " + sl.headline if sl.headline else "")
    h = max(0.32, text_h(head, 13, 7.0, 1.0) + 0.06) + 0.04  # 섹션 제목
    if sl.subline:
        h += 0.22
    blocks = [b for col in sl.columns for b in col]
    if blocks:
        h += 0.12 + sum(col_block_h(b) for b in blocks)
    h += 0.24 * len(sl.footnotes)
    if sl.bottom:
        h += 0.34
    return h + 0.14


def pill(s, x, y, w, h, t):
    b = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, I(x), I(y), I(w), I(h))
    _drop_style(b)
    b.fill.solid()
    b.fill.fore_color.rgb = LGRAY
    b.line.fill.background()
    tf = b.text_frame
    tf.margin_left = tf.margin_right = I(0.02)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.word_wrap = False
    p = tf.paragraphs[0]
    para_fmt(p, 1.0, align=PP_ALIGN.CENTER)
    size = 10.5 if text_w(t, 10.5) < w - 0.08 else 9
    runs(p, t, size, BLACK, True)


def col_section(s, sl: Slide, y: float, num: int | None):
    head = re.sub(r"^\[새쪽\]\s*", "", sl.chapter)
    head = re.sub(r"\s+_\s+", " : ", head)
    if num is not None and not re.match(r"^([0-9]+\.|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+\.|\[)", head):
        head = f"{num}. {head}"
    if sl.headline:
        head = f"{head} : {sl.headline}"
    hh = max(0.32, text_h(head, 13, 7.0, 1.0) + 0.06)
    tf = tbox(s, 0.19, y, 7.1, hh, MSO_ANCHOR.MIDDLE)
    para_fmt(tf.paragraphs[0], 1.0)
    runs(tf.paragraphs[0], head.replace("**", ""), 13, COL_KEY, True, key=COL_KEY)
    y += hh + 0.04
    if sl.subline:
        tf = tbox(s, 0.3, y - 0.04, 6.9, 0.24)
        runs(tf.paragraphs[0], sl.subline, 10, GRAY, key=COL_KEY)
        y += 0.22

    blocks = [b for col in sl.columns for b in col]
    if blocks:
        top = y
        y += 0.1
        for blk in blocks:
            x = COL_X + 0.11
            w = COL_BW - 0.22
            if blk.title and blk.lines and not blk.table:
                lh = col_lines_h(blk.lines, w - PILL_W - 0.1)
                py = y + (lh - 0.24) / 2 if lh < 0.45 else y + 0.02
                pill(s, x, py, PILL_W, 0.24, blk.title.strip("【】 "))
                col_draw_lines(s, x + PILL_W + 0.1, y, w - PILL_W - 0.1, blk.lines)
                y += max(lh, 0.28) + 0.08
                continue
            if blk.title:
                pill(s, x, y + 0.02, max(PILL_W, text_w(blk.title, 10.5) + 0.3), 0.24, blk.title.strip("【】 "))
                y += 0.3
            if blk.table:
                y += table(s, blk.table, x, y, w, 10.5, key=COL_KEY) + 0.08
            if blk.lines:
                y += col_draw_lines(s, x, y, w, blk.lines)
            y += 0.08
        outline_box(s, COL_X, top, COL_BW, y - top)
    for fn in sl.footnotes:
        tf = tbox(s, COL_X, y + 0.02, COL_BW, 0.22)
        runs(tf.paragraphs[0], fn, 9.5, GRAY, key=GRAY)
        y += 0.24
    if sl.bottom:
        tf = tbox(s, COL_X, y + 0.04, COL_BW, 0.3, MSO_ANCHOR.MIDDLE)
        p = tf.paragraphs[0]
        para_fmt(p, 1.0, align=PP_ALIGN.CENTER)
        runs(p, f"☞ {sl.bottom}", 12.5, COL_KEY, True, key=COL_KEY)
        y += 0.34
    return y + 0.14


def render_col(spec: Spec, template: Path, out: Path):
    prs = Presentation(template)
    drop_slides(prs, {0})
    first = prs.slides[0]
    layout = first.slide_layout
    keep = {"Picture 1", "TextBox 6"}      # 상단 배너 그림, 제목
    for sh in list(first.shapes):
        if sh.name not in keep:
            sh._element.getparent().remove(sh._element)
    title_box = next((sh for sh in first.shapes if sh.name == "TextBox 6"), None)
    if title_box is not None:
        set_para_text(title_box.text_frame.paragraphs[0], spec.meta.get("title", ""))
    stamp = " | ".join(v for v in (spec.meta.get("date", ""), spec.meta.get("dept", "")) if v)
    if stamp:
        tf = tbox(first, 4.4, 0.56, 3.0, 0.22)
        para_fmt(tf.paragraphs[0], 1.0, align=PP_ALIGN.RIGHT)
        runs(tf.paragraphs[0], stamp, 9, GRAY)

    def new_page():
        s = prs.slides.add_slide(layout)
        for ph in list(s.placeholders):
            ph._element.getparent().remove(ph._element)
        return s

    s, y, num = first, COL_TOP1, 0
    pages = 1
    for sl in spec.slides:
        if sl.kind == "section":          # 세로 양식은 간지 대신 새 쪽에서 시작
            if y > COL_TOPN + 0.01:
                s, y = new_page(), COL_TOPN
                pages += 1
            continue
        need = col_section_h(sl)
        force = sl.chapter.startswith("[새쪽]")
        if (force or y + need > COL_BOTTOM) and y > COL_TOPN + 0.01 and not (s is first and y == COL_TOP1):
            s, y = new_page(), COL_TOPN
            pages += 1
        if need > COL_BOTTOM - COL_TOPN:
            print(f"  ! '{sl.chapter}' 섹션: 한 쪽을 넘는 분량 — 줄이거나 나눌 것")
        if sl.kind == "body":
            num += 1
        y = col_section(s, sl, y, num if sl.kind == "body" else None)
        if sl.notes:
            s.notes_slide.notes_text_frame.text = "\n".join(sl.notes)
    prs.save(out)
    return pages


# ── 실행 ───────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser(description="report_spec.md → 회사 양식 PPTX (row 가로 / col A4 세로)")
    ap.add_argument("spec", type=Path)
    ap.add_argument("-o", "--out", type=Path)
    ap.add_argument("--layout", choices=("row", "col"))
    ap.add_argument("--template", type=Path, help="양식 파일을 직접 지정(기본: templates/row.pptx·col.pptx)")
    ap.add_argument("--check", action="store_true", help="양식 점검만 수행")
    args = ap.parse_args()

    spec = parse_spec(args.spec.read_text(encoding="utf-8-sig"))
    if not spec.slides:
        sys.exit("장표가 없습니다. '## 장 제목'으로 장표를 시작하세요.")
    msgs = check(spec)
    if args.check:
        print("\n".join(msgs) if msgs else "점검 통과: 지적 사항 없음")
        return
    layout = args.layout or spec.meta.get("layout", "row").strip().lower()
    if layout not in ("row", "col"):
        sys.exit("layout은 row 또는 col 입니다.")
    template = args.template or TEMPLATES / f"{layout}.pptx"
    if not template.exists():
        sys.exit(f"양식 파일이 없습니다: {template}")
    out = args.out or args.spec.with_name(args.spec.stem + f"_{layout}.pptx")
    if layout == "row":
        render_row(spec, template, out)
        n = sum(1 for sl in spec.slides if sl.kind == "body")
        print(f"생성: {out}  (가로 양식, 본문 {n}장)")
    else:
        pages = render_col(spec, template, out)
        print(f"생성: {out}  (A4 세로 양식, {pages}쪽)")
    if msgs:
        print("양식 점검 지적:")
        print("\n".join("  - " + m for m in msgs))


if __name__ == "__main__":
    main()
