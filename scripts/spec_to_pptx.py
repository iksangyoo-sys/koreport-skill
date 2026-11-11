#!/usr/bin/env python3
"""
koreport-skill — 보고서 문안(report_spec.md)을 1-page 장표 양식 PPTX로 렌더링한다.

LLM이 SVG를 장마다 직접 그리지 않고, 확정된 문안만 넘기면 이 스크립트가
결정론적으로 장표를 만든다. 토큰은 문안 작성에만 쓰인다.

사용법:
    python spec_to_pptx.py <report_spec.md> [-o out.pptx] [--check]

    --check   PPTX를 만들지 않고 양식 점검(줄 수·헤드라인 길이·기호 등)만 출력한다.

문법은 templates/report_spec_template.md 와 workflows/02-spec-format.md 참고.
의존성: python-pptx
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.oxml.ns import qn
    from pptx.util import Emu, Inches, Pt
except ImportError:  # pragma: no cover
    sys.exit("python-pptx가 필요합니다: pip install python-pptx")

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# ── 양식 상수 (koreport/skills/onepage-format §7) ─────────────────────────
FONT = "맑은 고딕"
C_TEXT = RGBColor(0x00, 0x00, 0x00)
C_HL = RGBColor(0x00, 0x33, 0xCC)       # 헤드라인·핵심어
C_DARK = RGBColor(0x00, 0x20, 0x60)     # 하단 따옴표·진한 제목
C_BAR = RGBColor(0x00, 0x70, 0xC0)      # 타이틀 바·표 헤더
C_FACE = RGBColor(0xD9, 0xEC, 0xF7)     # 보조 면
C_FACE2 = RGBColor(0xEE, 0xEE, 0xEE)
C_GRAY = RGBColor(0x80, 0x80, 0x80)
C_LINE = RGBColor(0xBF, 0xBF, 0xBF)
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)

SLIDE_W, SLIDE_H = Inches(13.333), Inches(7.5)
MARGIN = Inches(0.45)
CONTENT_W = SLIDE_W - 2 * MARGIN

LEVEL1 = ("❍", "■", "○", "□", "◎")
LEVEL2 = ("•",)
LEVEL3 = ("-", "–")
BODY_PREFIXES = LEVEL1 + LEVEL2 + LEVEL3 + ("☞", "※", "▷", "→", "①", "②", "③", "④", "⑤")


# ── 데이터 모델 ─────────────────────────────────────────────────────────
@dataclass
class Block:
    """본문 안의 한 덩어리: 【제목】 박스, 표, 또는 일반 문단 묶음."""
    title: str = ""
    lines: list[str] = field(default_factory=list)
    table: list[list[str]] = field(default_factory=list)


@dataclass
class Slide:
    chapter: str
    kind: str = "body"          # body | attach | section
    visual: bool = False        # @ppt-master 표시: PPT Master로 다시 그릴 장표
    headline: str = ""
    subline: str = ""
    columns: list[list[Block]] = field(default_factory=lambda: [[]])
    footnotes: list[str] = field(default_factory=list)
    bottom: str = ""
    notes: list[str] = field(default_factory=list)
    src_line: int = 0


@dataclass
class Spec:
    meta: dict[str, str]
    slides: list[Slide]


# ── 파서 ───────────────────────────────────────────────────────────────
def parse_spec(text: str) -> Spec:
    lines = text.replace("\r\n", "\n").split("\n")
    meta: dict[str, str] = {}
    i = 0
    if lines and lines[0].strip() == "---":
        i = 1
        while i < len(lines) and lines[i].strip() != "---":
            if ":" in lines[i]:
                k, v = lines[i].split(":", 1)
                meta[k.strip().lower()] = v.strip().strip('"')
            i += 1
        i += 1

    slides: list[Slide] = []
    cur: Slide | None = None
    in_comment = False

    def block() -> Block:
        col = cur.columns[-1]
        if not col or col[-1].table:
            col.append(Block())
        return col[-1]

    for n in range(i, len(lines)):
        raw = lines[n].rstrip()
        s = raw.strip()
        if in_comment:
            if "-->" in s:
                in_comment = False
            continue
        if s.startswith("<!--"):
            in_comment = "-->" not in s
            continue
        if s.startswith("## "):
            title = s[3:].strip()
            visual = "@ppt-master" in title
            title = title.replace("@ppt-master", "").strip()
            kind = "body"
            if re.match(r"^\[(첨부|참고)", title):
                kind = "attach"
            elif title.startswith("[간지]"):
                kind = "section"
                title = title[len("[간지]"):].strip()
            cur = Slide(chapter=title, kind=kind, visual=visual, src_line=n + 1)
            slides.append(cur)
            continue
        if cur is None or not s:
            continue
        if s.startswith("HL:"):
            cur.headline = s[3:].strip()
        elif s.startswith("SUB:"):
            cur.subline = s[4:].strip()
        elif s.startswith("NOTE:"):
            cur.notes.append(s[5:].strip())
        elif s == "||":
            cur.columns.append([])
        elif s.startswith("【"):
            cur.columns[-1].append(Block(title=s))
        elif s.startswith("|"):
            cells = [c.strip() for c in s.strip("|").split("|")]
            if all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c):
                continue
            col = cur.columns[-1]
            last = col[-1] if col else None
            if last is not None and (last.table or not last.lines):
                b = last  # 이어지는 표 행, 또는 【제목】 바로 아래 표
            else:
                b = Block()
                col.append(b)
            b.table.append(cells)
        elif s.startswith(("\"", "“")):
            cur.bottom = s.strip("\"“”")
        elif re.match(r"^\*{1,3}(\d+\))?\s", s) or re.match(r"^\d+\)\s", s):
            cur.footnotes.append(s)
        else:
            block().lines.append(s)
    return Spec(meta=meta, slides=slides)


# ── 텍스트 유틸 ─────────────────────────────────────────────────────────
def _set_font(run, size: float, color: RGBColor = C_TEXT, bold: bool = False, italic: bool = False):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = color
    f.name = FONT
    rpr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = rpr.makeelement(qn(tag), {})
            rpr.append(el)
        el.set("typeface", FONT)


def add_runs(paragraph, text: str, size: float, color: RGBColor = C_TEXT, bold: bool = False,
             italic: bool = False, emph: RGBColor = C_HL):
    """`**강조**` 구간은 블루 굵게 표시한다."""
    parts = re.split(r"(\*\*[^*]+\*\*)", text)
    for part in parts:
        if not part:
            continue
        r = paragraph.add_run()
        if part.startswith("**") and part.endswith("**"):
            r.text = part[2:-2]
            _set_font(r, size, emph, True, italic)
        else:
            r.text = part
            _set_font(r, size, color, bold, italic)


def visible_len(text: str) -> int:
    return len(re.sub(r"\*\*", "", text))


def est_lines(text: str, chars_per_line: int) -> int:
    # 한글 1자 ≈ 영문 2자 폭. 대략적인 줄 수 추정.
    w = sum(2 if ord(ch) > 0x2E80 else 1 for ch in re.sub(r"\*\*", "", text))
    return max(1, -(-w // (chars_per_line * 2)))


def line_style(s: str) -> dict:
    """기호에 따른 크기·색·들여쓰기."""
    if s.startswith(LEVEL1) or re.match(r"^[①-⑩❶-❿]", s):
        return dict(size=14, bold=True, indent=0, color=C_TEXT, space=8)
    if s.startswith(LEVEL2):
        return dict(size=13, bold=False, indent=1, color=C_TEXT, space=2)
    if s.startswith(LEVEL3):
        return dict(size=12, bold=False, indent=2, color=C_TEXT, space=1)
    if s.startswith("☞"):
        return dict(size=13.5, bold=True, indent=0, color=C_DARK, space=6)
    if s.startswith("※"):
        return dict(size=11, bold=False, indent=0, color=C_GRAY, space=2)
    if s.startswith("▷"):
        return dict(size=12.5, bold=False, indent=1, color=C_TEXT, space=2)
    return dict(size=13, bold=False, indent=0, color=C_TEXT, space=2)


# ── 도형 유틸 ───────────────────────────────────────────────────────────
def textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(int(x), int(y), int(w), int(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = Inches(0.03)
    return tb, tf


def rect(slide, x, y, w, h, fill: RGBColor | None, line: RGBColor | None = None, width_pt=0.75):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, int(x), int(y), int(w), int(h))
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(width_pt)
    shp.shadow.inherit = False
    return shp


def _drop_style(shape):
    st = shape._element.find(qn("p:style"))
    if st is not None:
        shape._element.remove(st)


def hline(slide, x, y, w, color=C_HL, width_pt=0.75):
    x, y, w = int(x), int(y), int(w)
    ln = slide.shapes.add_connector(1, x, y, x + w, y)
    _drop_style(ln)
    ln.line.color.rgb = color
    ln.line.width = Pt(width_pt)
    return ln


def _cell_border(cell, color="BFBFBF", width_emu=3175, sides=("T", "B")):
    # 스키마 순서: lnL, lnR, lnT, lnB 가 채우기(solidFill/noFill)보다 앞에 와야 한다.
    tcpr = cell._tc.get_or_add_tcPr()
    for pos, side in enumerate(("L", "R", "T", "B")):
        tag = qn(f"a:ln{side}")
        old = tcpr.find(tag)
        if old is not None:
            tcpr.remove(old)
        ln = tcpr.makeelement(tag, {"w": str(width_emu if side in sides else 0)})
        if side in sides:
            sf = ln.makeelement(qn("a:solidFill"), {})
            clr = sf.makeelement(qn("a:srgbClr"), {"val": color})
            sf.append(clr)
            ln.append(sf)
        else:
            ln.append(ln.makeelement(qn("a:noFill"), {}))
        tcpr.insert(pos, ln)


NUM_RE = re.compile(r"^[\s△▲▼↑↓+\-−]?[\d,.]+\s*(%p?|억|천|만|원|명|건|개|시간|배)?[\s↑↓]*$")


def add_table(slide, rows: list[list[str]], x, y, w, h_max):
    ncol = max(len(r) for r in rows)
    rows = [r + [""] * (ncol - len(r)) for r in rows]
    row_h = Inches(0.36)
    h = min(h_max, row_h * len(rows))
    gt = slide.shapes.add_table(len(rows), ncol, int(x), int(y), int(w), int(h))
    tbl = gt.table
    weights = [max(4, max(visible_len(r[ci]) * (2 if any(ord(ch) > 0x2E80 for ch in r[ci]) else 1)
                          for r in rows)) for ci in range(ncol)]
    for ci, wt in enumerate(weights):
        tbl.columns[ci].width = int(w * wt / sum(weights))
    # 기본 표 스타일 제거(밴드·외곽선 없애기)
    tblpr = tbl._tbl.tblPr
    tblpr.set("firstRow", "0")
    tblpr.set("bandRow", "0")
    for ri, r in enumerate(rows):
        for ci, val in enumerate(r):
            cell = tbl.cell(ri, ci)
            cell.margin_left = cell.margin_right = Inches(0.06)
            cell.margin_top = cell.margin_bottom = Inches(0.02)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            if ri == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = C_BAR
                add_runs(p, val, 12, C_WHITE, True, emph=C_WHITE)
                p.alignment = PP_ALIGN.CENTER
            else:
                cell.fill.background()
                add_runs(p, val, 12)
                p.alignment = PP_ALIGN.CENTER if NUM_RE.match(val or "x") else PP_ALIGN.LEFT
            _cell_border(cell)
    return h


# ── 렌더링 ─────────────────────────────────────────────────────────────
def render_cover(prs, meta: dict, toc: list[str]):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, SLIDE_W, Inches(0.18), C_BAR)
    _, tf = textbox(s, MARGIN, Inches(0.5), CONTENT_W, Inches(0.4))
    add_runs(tf.paragraphs[0], meta.get("date", ""), 14, C_GRAY)
    _, tf = textbox(s, Inches(1.2), Inches(2.2), SLIDE_W - Inches(2.4), Inches(1.6), MSO_ANCHOR.MIDDLE)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    add_runs(p, meta.get("title", "제목 없음"), 30, C_DARK, True)
    if meta.get("subtitle"):
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        add_runs(p2, meta["subtitle"], 16, C_HL)
    hline(s, Inches(3.2), Inches(3.95), SLIDE_W - Inches(6.4), C_BAR, 1.25)
    if toc:
        _, tf = textbox(s, Inches(2.5), Inches(4.25), SLIDE_W - Inches(5), Inches(0.5))
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        add_runs(p, "  ·  ".join(toc), 14, C_TEXT)
    _, tf = textbox(s, MARGIN, Inches(6.3), CONTENT_W, Inches(0.8), MSO_ANCHOR.BOTTOM)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    add_runs(p, meta.get("dept", ""), 16, C_TEXT, True)
    if meta.get("raci"):
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.RIGHT
        add_runs(p2, meta["raci"], 11, C_GRAY)


def render_section(prs, sl: Slide):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, Inches(3.1), SLIDE_W, Inches(1.3), C_BAR)
    _, tf = textbox(s, MARGIN, Inches(3.1), CONTENT_W, Inches(1.3), MSO_ANCHOR.MIDDLE)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    add_runs(p, sl.chapter, 28, C_WHITE, True, emph=C_WHITE)
    return s


def render_body(prs, sl: Slide, page: str):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    # ① 장 제목 바
    bar_h = Inches(0.55)
    rect(s, 0, 0, SLIDE_W, bar_h, C_BAR if sl.kind == "body" else C_DARK)
    _, tf = textbox(s, MARGIN, 0, CONTENT_W - Inches(1.2), bar_h, MSO_ANCHOR.MIDDLE)
    add_runs(tf.paragraphs[0], sl.chapter, 18, C_WHITE, True, emph=C_WHITE)
    _, tf = textbox(s, SLIDE_W - MARGIN - Inches(1.2), 0, Inches(1.2), bar_h, MSO_ANCHOR.MIDDLE)
    tf.paragraphs[0].alignment = PP_ALIGN.RIGHT
    add_runs(tf.paragraphs[0], page, 11, C_WHITE)

    y = bar_h + Inches(0.2)
    # ② 헤드라인
    if sl.headline:
        hl_lines = est_lines(sl.headline, 52)
        h = Inches(0.36) * hl_lines + Inches(0.1)
        _, tf = textbox(s, MARGIN, y, CONTENT_W, h, MSO_ANCHOR.MIDDLE)
        add_runs(tf.paragraphs[0], sl.headline, 17, C_HL, True, emph=C_HL)
        y += h + Inches(0.04)
        hline(s, MARGIN, y, CONTENT_W, C_HL, 0.75)
        y += Inches(0.1)
    if sl.subline:
        _, tf = textbox(s, MARGIN, y, CONTENT_W, Inches(0.36))
        add_runs(tf.paragraphs[0], sl.subline, 13.5, C_DARK)
        y += Inches(0.42)

    # 하단 영역 계산
    bottom_h = Inches(0.5) if sl.bottom else 0
    foot_h = Inches(0.24) * len(sl.footnotes)
    body_bottom = SLIDE_H - Inches(0.3) - bottom_h - foot_h
    gap = Inches(0.3)
    ncol = len(sl.columns)
    col_w = int((CONTENT_W - gap * (ncol - 1)) / ncol)
    y0 = y + Inches(0.08)

    for ci, col in enumerate(sl.columns):
        x = int(MARGIN + ci * (col_w + gap))
        cy = y0
        for blk in col:
            if blk.title:
                th = Inches(0.36)
                rect(s, x, cy, col_w, th, C_FACE)
                _, tf = textbox(s, x + Inches(0.08), cy, col_w - Inches(0.16), th, MSO_ANCHOR.MIDDLE)
                add_runs(tf.paragraphs[0], blk.title, 13, C_DARK, True)
                cy += th + Inches(0.08)
            if blk.table:
                th = add_table(s, blk.table, x, cy, col_w, body_bottom - cy)
                cy += th + Inches(0.15)
            if blk.lines:
                cpl = int(col_w / Inches(1) * 7.2)
                est = sum(est_lines(t, cpl) for t in blk.lines)
                h = min(body_bottom - cy, Inches(0.36) * est + Inches(0.14) * len(blk.lines))
                h = max(h, Inches(0.3))
                _, tf = textbox(s, x, cy, col_w, h)
                sparse = est <= 9
                first = True
                for t in blk.lines:
                    st = line_style(t)
                    p = tf.paragraphs[0] if first else tf.add_paragraph()
                    first_line = first
                    first = False
                    p.level = 0
                    p.space_before = Pt(st["space"] * (1.8 if sparse else 1) if not first_line else 0)
                    p.line_spacing = 1.15
                    if st["indent"]:
                        pPr = p._p.get_or_add_pPr()
                        mar = int(Inches(0.25) * st["indent"])
                        pPr.set("marL", str(mar + int(Inches(0.18))))
                        pPr.set("indent", str(-int(Inches(0.18))))
                    else:
                        pPr = p._p.get_or_add_pPr()
                        pPr.set("marL", str(int(Inches(0.22))))
                        pPr.set("indent", str(-int(Inches(0.22))))
                    add_runs(p, t, st["size"], st["color"], st["bold"])
                cy += h + Inches(0.12)
        if ncol > 1 and ci < ncol - 1:
            lx = int(x + col_w + gap / 2)
            ln = s.shapes.add_connector(1, lx, int(y0), lx, int(body_bottom))
            _drop_style(ln)
            ln.line.color.rgb = C_LINE
            ln.line.width = Pt(0.5)

    # * 각주
    fy = body_bottom + Inches(0.05)
    for fn in sl.footnotes:
        _, tf = textbox(s, MARGIN, fy, CONTENT_W, Inches(0.24))
        add_runs(tf.paragraphs[0], fn, 9.5, C_GRAY)
        fy += Inches(0.24)
    # ⑤ 하단 따옴표 문장
    if sl.bottom:
        by = SLIDE_H - Inches(0.25) - bottom_h
        hline(s, MARGIN, by, CONTENT_W, C_LINE, 0.5)
        _, tf = textbox(s, MARGIN, by + Inches(0.02), CONTENT_W, bottom_h, MSO_ANCHOR.MIDDLE)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        add_runs(p, f"“{sl.bottom}”", 13.5, C_DARK, True)
    if sl.visual:
        tag = rect(s, SLIDE_W - MARGIN - Inches(1.6), bar_h + Inches(0.02), Inches(1.6), Inches(0.22), None)
        tf = tag.text_frame
        tf.margin_top = tf.margin_bottom = 0
        add_runs(tf.paragraphs[0], "PPT Master 대상", 8, C_GRAY)
        tf.paragraphs[0].alignment = PP_ALIGN.RIGHT
    if sl.notes:
        s.notes_slide.notes_text_frame.text = "\n".join(sl.notes)
    return s


def render(spec: Spec, out: Path):
    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    toc = [t.strip() for t in spec.meta.get("toc", "").split("|") if t.strip()]
    if spec.meta.get("cover", "yes").lower() not in ("no", "false", "0"):
        render_cover(prs, spec.meta, toc)
    body = [sl for sl in spec.slides if sl.kind == "body"]
    total = len(body)
    k = 0
    for sl in spec.slides:
        if sl.kind == "section":
            render_section(prs, sl)
        elif sl.kind == "body":
            k += 1
            render_body(prs, sl, f"{k}/{total}")
        else:
            render_body(prs, sl, "")
    prs.save(out)


# ── 양식 점검 ───────────────────────────────────────────────────────────
def check(spec: Spec) -> list[str]:
    msgs: list[str] = []
    endings: dict[str, int] = {}
    for idx, sl in enumerate(spec.slides, 1):
        tag = f"[{idx}] {sl.chapter} (줄 {sl.src_line})"
        if sl.kind == "section":
            continue
        if sl.kind == "body" and not sl.headline:
            msgs.append(f"{tag}: 헤드라인(HL:)이 없음 — 결론 1~2줄을 넣을 것")
        if sl.headline and visible_len(sl.headline) > 100:
            msgs.append(f"{tag}: 헤드라인이 2줄을 넘을 가능성({visible_len(sl.headline)}자) — 키워드만 남길 것")
        if sl.headline and re.search(r"(검토|논의|고려)\s*(필요|예정)?\.?$", sl.headline):
            msgs.append(f"{tag}: 헤드라인이 '검토·논의·고려'로 끝남 — 무색무취한 결론(반려 원인)")
        lines = [t for col in sl.columns for b in col for t in b.lines]
        n_rows = sum(len(b.table) for col in sl.columns for b in col)
        total = len(lines) + n_rows + (1 if sl.headline else 0) + (1 if sl.subline else 0) + (1 if sl.bottom else 0)
        if len(sl.columns) > 1:
            total = max(
                sum(len(b.lines) + len(b.table) for b in col) for col in sl.columns
            ) + (1 if sl.headline else 0) + (1 if sl.subline else 0)
        if total > 15:
            msgs.append(f"{tag}: 약 {total}줄 — 13줄 안팎으로 줄이거나 장표를 나눌 것")
        tables = sum(1 for col in sl.columns for b in col if b.table)
        if tables > 2:
            msgs.append(f"{tag}: 표 {tables}개 — 장당 1~2개 권장")
        lv1 = [t for t in lines if t.startswith(LEVEL1)]
        if len({t[0] for t in lv1}) > 1:
            msgs.append(f"{tag}: 1단계 기호가 섞임({''.join(sorted({t[0] for t in lv1}))}) — 하나로 통일")
        per_col = max(sum(1 for b in col for t in b.lines if t.startswith(LEVEL1)) for col in sl.columns)
        if per_col > 3:
            msgs.append(f"{tag}: 한 단에 묶음 {per_col}개 — 3개 이하로 묶어 볼 것(3의 규칙)")
        for t in lines:
            if est_lines(t, 70) > 2:
                msgs.append(f"{tag}: 긴 불릿({visible_len(t)}자) — 1줄 목표, 3줄 넘으면 구조화: {t[:30]}…")
            if t.startswith("※") and t.rstrip().endswith(("필요", "해야 함")):
                msgs.append(f"{tag}: ※(보충)에 핵심 판단이 들어감 — ☞로 올릴 것: {t[:30]}…")
            m = re.search(r"(함|됨|임|음)\s*\.?$", t)
            if m:
                endings["서술형(~함·~됨)"] = endings.get("서술형(~함·~됨)", 0) + 1
        if re.search(r"현황|진단", sl.chapter):
            for t in lines:
                if re.search(r"필요\s*\.?$", t) and not t.startswith("☞"):
                    msgs.append(f"{tag}: 현황·진단 장에 '~필요'(권고) — 사실·이슈로 쓸 것: {t[:30]}…")
        for t in lines + [sl.headline]:
            if re.search(r"(TBD|XX|○○|확인\s*필요)", t or ""):
                msgs.append(f"{tag}: 미확정 표시가 남아 있음: {t[:40]}")
    if endings:
        msgs.append(f"[전체] 서술형 종결 {sum(endings.values())}줄 — 명사형 개조식과 섞였는지 확인")
    return msgs


# ── PPT Master 브리프 추출 ──────────────────────────────────────────────
def extract_visual(spec: Spec, src_text: str) -> str:
    """@ppt-master 장표만 뽑아 PPT Master Quick 입력용 브리프를 만든다."""
    blocks = re.split(r"(?m)^(?=## )", src_text)
    by_line = {}
    for b in blocks:
        if b.startswith("## "):
            by_line[b.splitlines()[0].strip()] = b.rstrip()
    cover = 1 if spec.meta.get("cover", "yes").lower() not in ("no", "false", "0") else 0
    out = [
        "# PPT Master 시각화 브리프",
        "",
        "아래 장표의 **문안은 확정본**이다. 헤드라인·본문 문구·수치를 바꾸지 말고 시각 구조(도형·도표·배치)만 설계한다.",
        "",
        "- 실행: PPT Master **Quick** 프로필(전략·확인 단계 생략). 장표 수 = 아래 목록 수, 순서 동일",
        "- 캔버스: 16:9 (1280×720), pptx_structure flat",
        "- 글꼴: 맑은 고딕. 헤드라인 16~17pt 굵게 #0033CC, 본문 12~14pt #000000, 각주 9~10pt",
        "- 색: 블루 단색 계열만(#0070C0 바·헤더, #002060 진한 제목, #D9ECF7·#C1DDF0 보조 면). 강조색은 핵심어에만, 3색 이하",
        "- 상단: 높이 0.55in 블루 바(#0070C0)에 흰색 장 제목 18pt, 우상단 쪽 번호 표기 그대로",
        "- 금지: 그라데이션·그림자·3D, 박스 안의 박스, 의미 없는 아이콘·스톡 이미지",
        "- 기호 위계(❍ • - ☞ ※ ▷ 【 】)를 그대로 유지",
        "",
        "| 기본 덱 쪽 | 시각화 덱 쪽 | 장 제목 |",
        "|---|---|---|",
    ]
    body_total = sum(1 for sl in spec.slides if sl.kind == "body")
    pos, k, maps, sections = cover, 0, [], []
    bk = 0
    for sl in spec.slides:
        pos += 1
        if sl.kind == "body":
            bk += 1
        if sl.visual:
            k += 1
            maps.append(f"{pos}={k}")
            out.append(f"| {pos} | {k} | {sl.chapter} |")
            src = next((v for key, v in by_line.items() if sl.chapter in key), "")
            sections.append(f"### 시각화 덱 {k}쪽 (쪽 번호 표기: {bk}/{body_total})\n\n{src}\n")
    if not k:
        return ""
    out += ["", f"병합 명령: `python scripts/merge_slides.py <기본.pptx> <ppt-master 결과.pptx> --map {','.join(maps)}`", ""]
    out += sections
    return "\n".join(out)


def main():
    ap = argparse.ArgumentParser(description="report_spec.md → 1-page 양식 PPTX")
    ap.add_argument("spec", type=Path)
    ap.add_argument("-o", "--out", type=Path)
    ap.add_argument("--check", action="store_true", help="양식 점검만 수행")
    ap.add_argument("--extract-visual", type=Path, metavar="BRIEF.md",
                    help="@ppt-master 장표만 뽑아 PPT Master 입력 브리프를 쓴다")
    args = ap.parse_args()

    src_text = args.spec.read_text(encoding="utf-8")
    spec = parse_spec(src_text)
    if not spec.slides:
        sys.exit("장표가 없습니다. '## 장 제목'으로 장표를 시작하세요.")
    msgs = check(spec)
    if args.check:
        print("\n".join(msgs) if msgs else "점검 통과: 지적 사항 없음")
        return
    if args.extract_visual:
        brief = extract_visual(spec, src_text)
        if not brief:
            sys.exit("@ppt-master 표시한 장표가 없습니다.")
        args.extract_visual.write_text(brief, encoding="utf-8")
        print(f"브리프: {args.extract_visual}")
        return
    out = args.out or args.spec.with_suffix(".pptx")
    render(spec, out)
    n_vis = sum(1 for sl in spec.slides if sl.visual)
    print(f"생성: {out}  (본문 {sum(1 for sl in spec.slides if sl.kind == 'body')}장"
          f"{', PPT Master 대상 ' + str(n_vis) + '장' if n_vis else ''})")
    if msgs:
        print("양식 점검 지적:")
        print("\n".join("  - " + m for m in msgs))


if __name__ == "__main__":
    main()
