#!/usr/bin/env python3
"""
koreport-skill — 문안을 회사 가로 양식(templates/row.pptx) '견본 배치' 그대로 꽉 찬 장표로 렌더링한다.

row.pptx 견본(본문 장)처럼 한 장 안에 아래 블록을 위에서부터 쌓고, 장 높이에 맞춰
글자 크기·간격을 자동으로 늘리거나 줄인다(빈 여백 없이 꽉 차게).

  ❍ 묶음 문장(15pt 굵게, **핵심어** #3333CC)
     * 각주(회색 10.5pt)
     ┌ 회색 0.25pt 선 박스 ──────────────────────────────┐
     │【라벨 바】 + 표       ┆  ((원형 라벨)) • 근거 문장   │   ← '||' 로 좌우 나눔
     └──────────────────────────────────────────────┘
     ☞ 시사점(13.5pt 굵게)
  “하단 따옴표 문장”(17pt 굵게 남색, 장 맨 아래 가운데)

사용법:
    python spec_to_onepage.py <onepage_spec.md> [-o 결과.pptx] [--pages N] [--template row.pptx]

    --pages N   만들 장 수. 문안에 '[새쪽]' 줄이 있으면 그 자리에서 나누고, 없으면
                ❍ 묶음 단위로 N장에 고르게 나눈다. 생략하면 [새쪽] 기준(없으면 1장),
                한 장에 안 들어가면 자동으로 다음 장으로 넘긴다.

문안 문법(onepage_spec.md) — workflows/02-spec-format.md '1-page 문법' 참고
    ---                         머리말: title(장 제목), page_titles(쪽별 제목, | 구분, 선택)
    ❍ 묶음 문장                 새 묶음(섹션). ■ ○ □ 도 같음
    * 각주 / ※ 보충             묶음 문장 바로 아래 회색 한 줄
    【라벨 제목】               박스 안 회색 라벨 바
    | 표 | 행 |                 표(첫 행 = 헤더). 셀 안 **값** 은 파란 굵게
    ((포항/집중))               원형 라벨('/'는 줄바꿈). ((인터뷰/결과|gray)) 는 회색
    • 근거 / - 하위 / ▷ 부연    박스 안 문장
    ||                          여기서부터 박스 오른쪽 단(가운데 점선)
    ☞ 시사점                    박스 아래 굵은 한 줄
    "문장" / “문장”             장 맨 아래 따옴표 문장(장마다 1개)
    [새쪽] 제목                 새 장(제목 생략 시 머리말 title)
    ## 장 제목 / HL: …          report_spec.md 형식도 읽는다(HL을 ❍ 묶음 문장으로 사용)

의존성: python-pptx
"""

from __future__ import annotations

import argparse
import math
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE.parent / "templates" / "row.pptx"
LAYOUT_NAME = "3_제목 및 내용"

FONT = "맑은 고딕"
BLACK = RGBColor(0, 0, 0)
NAVY = RGBColor(0x00, 0x20, 0x60)
KEY = RGBColor(0x33, 0x33, 0xCC)
GRAY = RGBColor(0x59, 0x59, 0x59)
LINEC = RGBColor(0xA6, 0xA6, 0xA6)

# 견본 장(row.pptx 3쪽) 좌표(인치)
HX, HW = 0.37, 10.1          # 묶음 문장
BX, BW = 0.56, 9.75          # 박스
TOP = 0.88                   # 제목 바 아래 시작
QUOTE_Y, QUOTE_H = 6.78, 0.42
BOTTOM_NOQUOTE = 7.12

LEVEL1 = ("❍", "■", "○", "□", "◎")


def I(v: float) -> int:
    return int(Inches(v))


# ── 문안 해석 ─────────────────────────────────────────────────────────
@dataclass
class Item:
    kind: str                 # label | table | circle | text
    text: str = ""
    rows: list = field(default_factory=list)
    style: str = ""           # text: bullet/sub/arrow/small/plain, circle: 색


@dataclass
class Section:
    heading: str = ""
    notes: list = field(default_factory=list)
    cols: list = field(default_factory=lambda: [[]])
    split: float | None = None   # 왼쪽 단 비율
    after: list = field(default_factory=list)   # ☞ 줄


@dataclass
class Page:
    title: str = ""
    sections: list = field(default_factory=list)
    quote: str = ""


def _front(text: str):
    meta = {}
    m = re.match(r"^---\s*\n(.*?)\n---\s*\n", text, re.S)
    if m:
        for ln in m.group(1).splitlines():
            if ":" in ln and not ln.strip().startswith("#"):
                k, v = ln.split(":", 1)
                meta[k.strip()] = v.strip().strip("'\"") if v.strip()[:1] in "'\"" and v.strip()[-1:] in "'\"" else v.strip()
        text = text[m.end():]
    return meta, text


def parse(path: Path):
    meta, body = _front(path.read_text(encoding="utf-8"))
    body = re.sub(r"<!--.*?-->", "", body, flags=re.S)
    title = meta.get("title", path.stem)
    pages = [Page(title=title)]
    explicit_break = False
    sec: Section | None = None
    table: list | None = None

    def cur_col():
        return sec.cols[-1]

    def ensure_sec():
        nonlocal sec
        if sec is None:
            sec = Section()
            pages[-1].sections.append(sec)
        return sec

    def flush_table():
        nonlocal table
        if table:
            ensure_sec()
            cur_col().append(Item("table", rows=table))
        table = None

    for raw in body.splitlines():
        ln = raw.strip()
        if not ln:
            flush_table()
            continue
        if ln.startswith("|") and ln != "||" and not re.fullmatch(r"\|\|\s*\d+", ln):
            cells = [c.strip() for c in ln.strip("|").split("|")]
            if all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c):
                continue
            table = (table or []) + [cells]
            continue
        flush_table()
        if ln.startswith("[새쪽]") or ln.startswith("## [새쪽]"):
            t = re.sub(r"^(##\s*)?\[새쪽\]\s*", "", ln)
            pages.append(Page(title=t or title))
            explicit_break = True
            sec = None
            continue
        if ln.startswith("# ") and not ln.startswith("## "):
            pages[-1].title = ln[2:].strip() if len(pages) > 1 or title == path.stem else pages[-1].title
            continue
        if ln.startswith("## "):            # report_spec.md 형식: 장 = 묶음
            sec = Section(heading="")
            sec._chapter = re.sub(r"\s*@ppt-master\s*$", "", ln[3:].strip())  # type: ignore[attr-defined]
            pages[-1].sections.append(sec)
            continue
        if ln.startswith("HL:"):
            ensure_sec().heading = "❍ " + ln[3:].strip()
            continue
        if ln.startswith(("SUB:", "NOTE:")):
            continue
        if ln.startswith(LEVEL1):
            if sec is not None and getattr(sec, "_chapter", None) is not None and sec.heading:
                # 장 안의 ❍ 줄은 굵은 근거 문장으로
                cur_col().append(Item("text", "__" + ln[1:].strip() + "__", style="bullet"))
                continue
            if sec is not None and getattr(sec, "_chapter", None) is not None and not sec.heading:
                sec.heading = "❍ " + ln[1:].strip()
                continue
            sec = Section(heading="❍ " + ln[1:].strip())
            pages[-1].sections.append(sec)
            continue
        m = re.fullmatch(r"\|\|\s*(\d+)?", ln)
        if m:
            s = ensure_sec()
            s.cols.append([])
            if m.group(1):
                s.split = int(m.group(1)) / 100
            continue
        if (ln[0] in "\"“") and (ln[-1] in "\"”") and len(ln) > 2:
            pages[-1].quote = ln.strip("\"“”")
            continue
        if ln.startswith("【") and ln.endswith("】"):
            ensure_sec(); cur_col().append(Item("label", ln[1:-1].strip()))
            continue
        m = re.fullmatch(r"\(\((.+?)(?:\|(\w+))?\)\)", ln)
        if m:
            ensure_sec(); cur_col().append(Item("circle", m.group(1).replace("/", "\n"), style=m.group(2) or ""))
            continue
        if ln.startswith("☞"):
            ensure_sec().after.append(ln)
            continue
        s = ensure_sec()
        if ln.startswith(("* ", "※", "** ", "1)")) and not any(s.cols) :
            s.notes.append(ln)
            continue
        if ln.startswith("•"):
            cur_col().append(Item("text", ln, style="bullet"))
        elif ln.startswith("- "):
            cur_col().append(Item("text", ln, style="sub"))
        elif ln.startswith(("▷", "→")):
            cur_col().append(Item("text", ln, style="arrow"))
        elif ln.startswith(("※", "* ")):
            cur_col().append(Item("text", ln, style="small"))
        else:
            cur_col().append(Item("text", ln, style="plain"))
    flush_table()
    # 장 이름만 있고 HL이 없는 장
    for p in pages:
        for s in p.sections:
            if not s.heading and getattr(s, "_chapter", None):
                s.heading = "❍ " + re.sub(r"^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩIVX\d]+\.\s*", "", s._chapter)  # type: ignore[attr-defined]
    pages = [p for p in pages if p.sections or p.quote]
    return meta, pages, explicit_break


# ── 크기 추정 ─────────────────────────────────────────────────────────
def _plain(t: str) -> str:
    return re.sub(r"\*\*|__", "", t)


def text_w(t: str, size: float) -> float:
    w = 0.0
    for ch in _plain(t):
        o = ord(ch)
        w += 1.0 if (o >= 0x1100 and not 0x2000 <= o <= 0x206F) else (0.3 if ch == " " else 0.56)
    return w * size / 72


def n_lines(t: str, size: float, width: float) -> int:
    return max(1, math.ceil(text_w(t, size) * 1.06 / max(width - 0.1, 0.3)))


def lh(size: float, spacing: float = 1.2) -> float:
    return size * spacing * 1.2 / 72


TEXT_SIZE = {"bullet": 12.5, "sub": 12, "arrow": 12, "small": 10.5, "plain": 12.5}
TEXT_HANG = {"bullet": 0.16, "sub": 0.3, "arrow": 0.16, "small": 0.0, "plain": 0.0}


def col_width_weights(rows, size):
    nc = max(len(r) for r in rows)
    ws = [0.0] * nc
    for r in rows:
        for j, c in enumerate(r):
            ws[j] = max(ws[j], text_w(c, size) + 0.2)
    return ws


def table_geom(rows, width, k):
    size = 10.5 * k
    ws = col_width_weights(rows, size)
    tot = sum(ws)
    cw = [width * w / tot for w in ws]
    heights = []
    for r in rows:
        nl = max(n_lines(c, size, cw[j] - 0.08) for j, c in enumerate(r))
        heights.append(max(0.235 * k, nl * size * 1.2 / 72 + 0.07))
    return cw, heights


def item_h(it: Item, width: float, k: float) -> float:
    if it.kind == "label":
        return 0.3 * k + 0.06
    if it.kind == "table":
        return sum(table_geom(it.rows, width, k)[1]) + 0.08
    if it.kind == "text":
        size = TEXT_SIZE[it.style] * k
        return n_lines(it.text, size, width - TEXT_HANG[it.style] - 0.1) * lh(size, 1.15) + 3 / 72
    return 0.0


def col_layout(items, width, k):
    """원형 라벨이 있으면 왼쪽에 두고 나머지를 오른쪽에 쌓는다. (높이, 원 지름) 반환."""
    circ = [it for it in items if it.kind == "circle"]
    d = 0.66 * min(k, 1.15) if circ else 0.0
    inner = width - (d + 0.3 if circ else 0)
    h = sum(item_h(it, inner, k) for it in items if it.kind != "circle")
    return max(h, d + 0.12 * len(circ)), d, inner


def split_ratio(sec: Section) -> float:
    if sec.split:
        return sec.split
    left, right = sec.cols[0], sec.cols[1]
    if any(i.kind == "table" for i in left) and not any(i.kind == "table" for i in right):
        return 0.57
    if any(i.kind == "table" for i in right) and not any(i.kind == "table" for i in left):
        return 0.43
    return 0.5


def col_widths(sec: Section):
    if len(sec.cols) == 1:
        return [BW - 0.2]
    r = split_ratio(sec)
    lw = BW * r
    return [lw - 0.2, BW - lw - 0.2]


def heading_h(sec, k):
    size = 15 * k
    return n_lines(sec.heading, size, HW) * lh(size, 1.0) + 0.06


def notes_h(sec, k):
    size = 10.5 * k
    return sum(n_lines(n, size, BW) * lh(size, 1.0) for n in sec.notes)


def has_box(sec):
    return any(sec.cols) and (len(sec.cols) > 1 or any(i.kind != "text" for c in sec.cols for i in c))


def body_h(sec, k):
    if not any(sec.cols):
        return 0.0
    hs = [col_layout(c, w, k)[0] for c, w in zip(sec.cols, col_widths(sec))]
    return max(hs) + (0.2 if has_box(sec) else 0.04)


def after_h(sec, k):
    size = 13.5 * k
    return sum(n_lines(a, size, BW) * lh(size, 1.0) + 0.04 for a in sec.after)


def section_h(sec, k):
    return heading_h(sec, k) + notes_h(sec, k) + 0.04 + body_h(sec, k) + after_h(sec, k)


def avail(page: Page) -> float:
    return (QUOTE_Y - 0.06 if page.quote else BOTTOM_NOQUOTE) - TOP


SCALES = [1.15, 1.1, 1.05, 1.0, 0.96, 0.92, 0.88, 0.85, 0.82]
GAP = 0.12


def fit(page: Page):
    """장이 꽉 차는 가장 큰 배율과 남는 높이를 돌려준다. 넘치면 (None, 초과량)."""
    n = len(page.sections)
    for k in SCALES:
        need = sum(section_h(s, k) for s in page.sections) + GAP * (n - 1)
        if need <= avail(page):
            return k, avail(page) - need
    k = SCALES[-1]
    return None, sum(section_h(s, k) for s in page.sections) + GAP * (n - 1) - avail(page)


# ── 그리기 유틸 (견본 규격) ───────────────────────────────────────────
def runs(p, text, size, color=BLACK, bold=False):
    for part in re.split(r"(\*\*[^*]+\*\*|__[^_]+__)", text):
        if not part:
            continue
        c, b = color, bold
        if part.startswith("**"):
            part, c, b = part[2:-2], KEY, True
        elif part.startswith("__"):
            part, b = part[2:-2], True
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


def para_fmt(p, line=1.15, before=0, align=PP_ALIGN.LEFT, hang=0.0):
    p.line_spacing = line
    p.space_before = Pt(before)
    p.alignment = align
    if hang:
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", str(I(hang)))
        pPr.set("indent", str(-I(hang)))


def tbox(s, x, y, w, h, anchor=MSO_ANCHOR.TOP, inset=0.05):
    tb = s.shapes.add_textbox(I(x), I(y), I(w), I(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = I(inset)
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
    b.line.color.rgb = LINEC
    b.line.width = Pt(0.25)


def patt_fill(shp, prst, fg_hex):
    spPr = shp._element.spPr
    for tag in ("a:solidFill", "a:noFill", "a:pattFill", "a:gradFill"):
        for el in spPr.findall(qn(tag)):
            spPr.remove(el)
    patt = etree.SubElement(spPr, qn("a:pattFill"), prst=prst)
    fg = etree.SubElement(patt, qn("a:fgClr"))
    etree.SubElement(fg, qn("a:srgbClr"), val=fg_hex)
    bg = etree.SubElement(patt, qn("a:bgClr"))
    etree.SubElement(bg, qn("a:srgbClr"), val="FFFFFF")
    ln = spPr.find(qn("a:ln"))
    if ln is not None:
        spPr.remove(patt)
        ln.addprevious(patt)


def label_bar(s, x, y, w, h, text, size):
    b = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, I(x), I(y), I(w), I(h))
    _drop_style(b)
    patt_fill(b, "pct90", "D9D9D9")
    b.line.fill.background()
    tf = b.text_frame
    tf.margin_left, tf.margin_right = I(0.1), I(0.05)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    para_fmt(p, 1.0, align=PP_ALIGN.CENTER)
    runs(p, text, size, NAVY, True)


CIRCLE_COLOR = {"": "B3F3E5", "mint": "B3F3E5", "gray": "D9D9D9", "grey": "D9D9D9",
                "blue": "BDD7EE", "yellow": "FFE699"}


def circle(s, cx, cy, d, text, style, size):
    c = s.shapes.add_shape(MSO_SHAPE.OVAL, I(cx - d / 2), I(cy - d / 2), I(d), I(d))
    _drop_style(c)
    patt_fill(c, "pct70", CIRCLE_COLOR.get(style.lower(), "B3F3E5"))
    c.line.fill.background()
    tf = c.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.word_wrap = False
    for i, t in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para_fmt(p, 1.0, align=PP_ALIGN.CENTER)
        runs(p, t, size, NAVY, True)


def dashed_v(s, x, y1, y2):
    ln = s.shapes.add_connector(1, I(x), I(y1), I(x), I(y2))
    _drop_style(ln)
    ln.line.color.rgb = LINEC
    ln.line.width = Pt(0.5)
    etree.SubElement(ln.line._get_or_add_ln(), qn("a:prstDash"), val="sysDot")


def _border(cell, side, w, color="A6A6A6", none=False):
    tcPr = cell._tc.get_or_add_tcPr()
    tag = {"L": "a:lnL", "R": "a:lnR", "T": "a:lnT", "B": "a:lnB"}[side]
    for el in tcPr.findall(qn(tag)):
        tcPr.remove(el)
    ln = etree.SubElement(tcPr, qn(tag), w="0" if none else str(int(w)))
    if none:
        etree.SubElement(ln, qn("a:noFill"))
    else:
        sf = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr"), val=color)
    for t in ("a:lnL", "a:lnR", "a:lnT", "a:lnB", "a:solidFill", "a:noFill"):
        for el in tcPr.findall(qn(t)):
            tcPr.remove(el)
            tcPr.append(el)


def draw_table(s, x, y, w, rows, k):
    size = 10.5 * k
    cw, hs = table_geom(rows, w, k)
    nc = len(cw)
    rows = [r + [""] * (nc - len(r)) for r in rows]
    nr = len(rows)
    gt = s.shapes.add_table(nr, nc, I(x), I(y), I(w), I(sum(hs)))
    tbl = gt.table
    tblPr = tbl._tbl.tblPr
    for a in ("firstRow", "bandRow"):
        tblPr.set(a, "0")
    st = tblPr.find(qn("a:tableStyleId"))
    if st is not None:
        st.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"
    for j, c in enumerate(cw):
        tbl.columns[j].width = I(c)
    for i, row in enumerate(rows):
        tbl.rows[i].height = I(hs[i])
        hdr = i == 0
        for j, val in enumerate(row):
            c = tbl.cell(i, j)
            c.margin_left = c.margin_right = I(0.04)
            c.margin_top = c.margin_bottom = I(0.01)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = c.text_frame.paragraphs[0]
            para_fmt(p, 1.0, align=PP_ALIGN.LEFT if (j == 0 and not hdr) else PP_ALIGN.CENTER)
            runs(p, val, size, NAVY if hdr else BLACK, hdr)
            if hdr:
                c.fill.solid()
                c.fill.fore_color.rgb = RGBColor(0xF2, 0xF2, 0xF2)
            else:
                c.fill.background()
            _border(c, "L", 0, none=True)
            _border(c, "R", 0, none=True)
            _border(c, "T", Pt(1) if hdr else Pt(0.25), "7F7F7F" if hdr else "BFBFBF")
            _border(c, "B", Pt(1) if i == nr - 1 else Pt(0.25), "7F7F7F" if i == nr - 1 else "BFBFBF")
    return sum(hs)


def draw_col(s, items, x, y, width, height, k):
    h, d, inner = col_layout(items, width, k)
    circs = [it for it in items if it.kind == "circle"]
    if circs:
        n = len(circs)
        for i, c in enumerate(circs):
            cy = y + height * (i + 0.5) / n
            circle(s, x + 0.05 + d / 2, cy, d, c.text, c.style, 11 * min(k, 1.1))
        x = x + d + 0.3
    # 내용을 세로 가운데로
    content_h = sum(item_h(it, inner, k) for it in items if it.kind != "circle")
    cy = y + max(0.0, (height - content_h) / 2) if circs else y
    run_text: list = []

    def flush(yy):
        if not run_text:
            return yy
        tf = tbox(s, x, yy, inner, sum(item_h(t, inner, k) for t in run_text))
        for i, t in enumerate(run_text):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            size = TEXT_SIZE[t.style] * k
            para_fmt(p, 1.15, 0 if i == 0 else 3, hang=TEXT_HANG[t.style])
            runs(p, t.text, size, GRAY if t.style == "small" else BLACK)
        yy += sum(item_h(t, inner, k) for t in run_text)
        run_text.clear()
        return yy

    for it in items:
        if it.kind == "circle":
            continue
        if it.kind == "text":
            run_text.append(it)
            continue
        cy = flush(cy)
        if it.kind == "label":
            label_bar(s, x - 0.1 if not circs else x, cy, inner + (0.2 if not circs else 0), 0.3 * k, it.text, 11.5 * k)
            cy += 0.3 * k + 0.06
        elif it.kind == "table":
            cy += draw_table(s, x, cy, inner, it.rows, k) + 0.08
    flush(cy)


def draw_section(s, sec: Section, y, k, extra):
    size = 15 * k
    hh = heading_h(sec, k)
    tf = tbox(s, HX, y, HW, hh)
    p = tf.paragraphs[0]
    para_fmt(p, 1.0, hang=0.25)
    runs(p, sec.heading, size, bold=True)
    y += hh
    for n in sec.notes:
        nh = n_lines(n, 10.5 * k, BW) * lh(10.5 * k, 1.0)
        tf = tbox(s, BX, y, BW, nh)
        runs(tf.paragraphs[0], n, 10.5 * k, GRAY)
        y += nh
    y += 0.04
    if any(sec.cols):
        bh = body_h(sec, k) + extra
        widths = col_widths(sec)
        if has_box(sec):
            outline_box(s, BX, y, BW, bh)
            xs = [BX + 0.1]
            if len(sec.cols) > 1:
                lw = BW * split_ratio(sec)
                dashed_v(s, BX + lw, y + 0.1, y + bh - 0.1)
                xs.append(BX + lw + 0.1)
            for c, x, w in zip(sec.cols, xs, widths):
                # 왼쪽 단의 라벨 바는 박스 위 가장자리에 붙인다(견본)
                top_pad = 0.0 if (c and c[0].kind == "label") else 0.1
                draw_col(s, c, x, y + top_pad, w, bh - top_pad - 0.1, k)
        else:
            draw_col(s, sec.cols[0], BX + 0.1, y, widths[0], bh, k)
        y += bh
    for a in sec.after:
        ah = n_lines(a, 13.5 * k, BW) * lh(13.5 * k, 1.0) + 0.04
        tf = tbox(s, BX, y + 0.02, BW, ah)
        runs(tf.paragraphs[0], a, 13.5 * k, bold=True)
        y += ah
    return y


# ── 장 나누기 ─────────────────────────────────────────────────────────
def paginate(pages, want, explicit):
    if want and not explicit and want > 1 and len(pages) == 1:
        p0 = pages[0]
        secs = p0.sections
        tot = [section_h(x, 1.0) for x in secs]
        target = sum(tot) / want
        out, cur, acc = [], [], 0.0
        for sct, h in zip(secs, tot):
            if cur and acc + h / 2 > target * (len(out) + 1) and len(out) < want - 1:
                out.append(cur)
                cur = []
            cur.append(sct)
            acc += h
        out.append(cur)
        pages = [Page(title=p0.title, sections=c) for c in out]
        pages[-1].quote = p0.quote
    # 넘치는 장은 묶음 단위로 다음 장으로 넘긴다
    res = []
    for p in pages:
        queue = list(p.sections)
        while queue:
            np_ = Page(title=p.title, quote="")
            while queue:
                np_.sections.append(queue.pop(0))
                np_.quote = p.quote if not queue else ""
                if fit(np_)[0] is None and len(np_.sections) > 1:
                    queue.insert(0, np_.sections.pop())
                    np_.quote = ""
                    break
            res.append(np_)
    return res


def build(spec: Path, out: Path, template: Path, want: int | None):
    meta, pages, explicit = parse(spec)
    if not pages:
        sys.exit("문안에서 ❍ 묶음을 찾지 못했습니다.")
    if meta.get("pages") and not want:
        want = int(meta["pages"])
    pages = paginate(pages, want, explicit)
    titles = [t.strip() for t in meta.get("page_titles", "").split("|") if t.strip()]
    for i, p in enumerate(pages):
        if i < len(titles):
            p.title = titles[i]

    prs = Presentation(str(template))
    layout = prs.slides[2].slide_layout if len(prs.slides) > 2 else None
    if layout is None or layout.name != LAYOUT_NAME:
        layout = next((l for l in prs.slide_layouts if l.name == LAYOUT_NAME), prs.slide_layouts[0])
    lst = prs.slides._sldIdLst
    for sld in list(lst):
        prs.part.drop_rel(sld.get(qn("r:id")))
        lst.remove(sld)

    warns = []
    N = len(pages)
    for no, pg in enumerate(pages, 1):
        s = prs.slides.add_slide(layout)
        for ph in list(s.placeholders):
            if ph.placeholder_format.idx != 0:
                ph._element.getparent().remove(ph._element)
        if s.shapes.title is not None:
            s.shapes.title.text_frame.text = pg.title
        tf = tbox(s, 9.96, 0.31, 0.79, 0.4, MSO_ANCHOR.MIDDLE)
        p = tf.paragraphs[0]
        para_fmt(p, 1.0, align=PP_ALIGN.CENTER)
        runs(p, f"__{no}__/{N}", 13, GRAY)

        k, spare = fit(pg)
        if k is None:
            warns.append(f"{no}쪽: 약 {spare:.2f}인치 넘칠 수 있음 → 문안을 줄이거나 장 수를 늘리세요")
            k, spare = SCALES[-1], 0.0
        n = len(pg.sections)
        boxes = [i for i, sc in enumerate(pg.sections) if any(sc.cols)]
        # 남는 높이: 묶음 사이 간격(최대 0.15)과 박스 높이에 나눠 꽉 채운다
        gap_add = min(0.15, spare * 0.4 / max(n - 1, 1)) if n > 1 else 0.0
        rest = spare - gap_add * (n - 1)
        box_add = min(0.45, rest / len(boxes)) if boxes else 0.0  # 너무 크면 빈 박스가 되므로 상한
        y = TOP
        for i, sec in enumerate(pg.sections):
            y = draw_section(s, sec, y, k, box_add if i in boxes else 0.0)
            y += GAP + gap_add
        if pg.quote:
            tf = tbox(s, 0.0, QUOTE_Y, 10.83, QUOTE_H, MSO_ANCHOR.MIDDLE)
            p = tf.paragraphs[0]
            para_fmt(p, 1.0, align=PP_ALIGN.CENTER)
            q = 17 if text_w(pg.quote, 17) < 10.4 else 15
            runs(p, f"“{pg.quote}”", q, NAVY, True)
        print(f"  {no}쪽 '{pg.title}': 묶음 {n}개, 글자 배율 {k:.2f}")
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    print("생성:", out)
    for w in warns:
        print("  ⚠", w)


def main():
    ap = argparse.ArgumentParser(description="문안 → row.pptx 견본 배치 1-page 장표")
    ap.add_argument("spec")
    ap.add_argument("-o", "--out")
    ap.add_argument("--pages", type=int, help="만들 장 수(기본: [새쪽] 기준, 없으면 1장)")
    ap.add_argument("--template", default=str(TEMPLATE))
    a = ap.parse_args()
    spec = Path(a.spec)
    meta, _ = _front(spec.read_text(encoding="utf-8"))
    out = Path(a.out) if a.out else spec.with_name(re.sub(r'[\\/:*?"<>|]', "", meta.get("title", spec.stem)) + ".pptx")
    tpl = Path(a.template)
    if not tpl.exists():
        sys.exit(f"양식 파일이 없습니다: {tpl}  (templates/row.pptx를 넣거나 --template으로 지정)")
    build(spec, out, tpl, a.pages)


if __name__ == "__main__":
    main()
