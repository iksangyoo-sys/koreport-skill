#!/usr/bin/env python3
"""
koreport-skill — 안내문·공지 문안(notice_spec.md)을 '텍스트 중심' 1-page 장표로 렌더링한다.

안내문은 독자가 '찾아 읽는' 문서다. 그래서 ❍ 묶음마다 박스·라벨 바·원형 라벨을 두는
spec_to_onepage.py(보고서용) 대신, 아래처럼 위에서 아래로 훑어 내려가는 배치를 쓴다.

  ┌──────────────────────────────────────────────┐
  │  핵심 메시지 한 문장 (회색 테두리 박스, 가운데)   │   ← '> 문장'
  └──────────────────────────────────────────────┘
  ●  교육 대상: 값                                      ← '● 라벨: 값'  (라벨 굵게)
       * 각주                                           ← '* ' / '** ' (회색 작은 글씨)
  ●  교육 기간: **핵심 값** → 세부                       ← **값** = 파란 굵게, __값__ = 검정 굵게
       ▶ 수료 기준: 출석 80% 이상     ※ 오른쪽 보충       ← '▶ … >> ※ …' (오른쪽 정렬 보충)
  ┌ 표(일정처럼 2차원 정보만) ────────────────────┐   ← '| … |' (첫 행 헤더=하늘색, 첫 열=회색)
  └──────────────────────────────────────────────┘
       - 하위 항목                                      ← '- '
  ─────────────────────────────────────────────  ← '---' (본문 아래 구분선)
  ■  문의: …                                           ← '■ …' (맨 아래 꼬리 줄)

사용법:
    python spec_to_notice.py <notice_spec.md> [-o 결과.pptx] [--template row.pptx]

    --template  회사 양식(templates/row.pptx 등). 있으면 그 마스터를 쓰고 자리표시자는 지운다.
                없으면 10.83x7.5in 빈 장에 그린다. 글자 크기는 장 높이에 맞춰 자동 조정한다.

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

FONT = "맑은 고딕"
BLACK = RGBColor(0, 0, 0)
NAVY = RGBColor(0x1F, 0x2F, 0x5C)
BLUE = RGBColor(0x00, 0x20, 0xE0)
GRAY = RGBColor(0x40, 0x40, 0x40)
W_IN, H_IN = 10.83, 7.5
MX, TOP, BOTTOM = 0.3, 0.28, 7.25

# 기본 글자 크기(pt) — k 배율로 함께 늘고 준다
SIZE = {"msg": 19, "l1": 15.5, "sub": 14, "foot": 11, "table": 12, "tail": 13.5}
INDENT = {"l1": 0.0, "sub": 0.42, "foot": 0.45}


@dataclass
class Blk:
    kind: str                 # msg | l1 | sub | foot | table | hr | tail
    text: str = ""
    right: str = ""           # 같은 줄 오른쪽 보충(>>)
    rows: list = field(default_factory=list)


def _front(text: str):
    meta = {}
    if text.startswith("---"):
        end = text.find("\n---", 3)
        if end > 0:
            for ln in text[3:end].splitlines():
                if ":" in ln:
                    k, v = ln.split(":", 1)
                    meta[k.strip()] = v.strip()
            text = text[end + 4:]
    return meta, text


def parse(path: Path):
    meta, body = _front(path.read_text(encoding="utf-8"))
    body = re.sub(r"<!--.*?-->", "", body, flags=re.S)
    blks: list[Blk] = []
    table: list = []

    def flush():
        nonlocal table
        if table:
            blks.append(Blk("table", rows=table))
            table = []

    for raw in body.splitlines():
        if not raw.strip():
            flush()
            continue
        ln = raw.strip()
        if ln.startswith("|"):
            cells = [c.strip() for c in ln.strip("|").split("|")]
            if all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c):
                continue
            table.append(cells)
            continue
        flush()
        right = ""
        if ">>" in ln:
            ln, right = [t.strip() for t in ln.split(">>", 1)]
        if ln.startswith(">"):
            blks.append(Blk("msg", ln[1:].strip()))
        elif re.fullmatch(r"-{3,}|─{3,}", ln):
            blks.append(Blk("hr"))
        elif ln.startswith("■"):
            blks.append(Blk("tail", ln[1:].strip()))
        elif ln[0] in "●❍○◎":
            blks.append(Blk("l1", ln[1:].strip(), right))
        elif ln.startswith("▶") or ln.startswith("- ") or ln.startswith("※"):
            blks.append(Blk("sub", ln, right))
        elif ln.startswith("*"):
            blks.append(Blk("foot", ln, right))
        else:
            blks.append(Blk("sub" if raw[:1] in " \t" else "l1", ln, right))
    flush()
    return meta, blks


# ── 크기 추정 ─────────────────────────────────────────────
def _plain(t):
    return re.sub(r"\*\*|__", "", t)


def text_w(t, size):
    w = 0.0
    for ch in _plain(t):
        o = ord(ch)
        w += 1.0 if (o >= 0x1100 and not 0x2000 <= o <= 0x206F) else (0.3 if ch == " " else 0.56)
    return w * size / 72


def n_lines(t, size, width):
    return max(1, math.ceil(text_w(t, size) * 1.12 / max(width, 0.5)))


def line_h(size):
    return size * 1.25 / 72


def table_geom(rows, width, k):
    size = SIZE["table"] * k
    nc = max(len(r) for r in rows)
    first = max(text_w(r[0], size) for r in rows) + 0.25
    first = min(max(first, width * 0.14), width * 0.3)
    rest = (width - first) / max(nc - 1, 1)
    cw = [first] + [rest] * (nc - 1)
    hs = []
    for r in rows:
        nl = max(n_lines(c, size, cw[j] - 0.1) for j, c in enumerate(r + [""] * (nc - len(r))))
        hs.append(max(0.42 * k, nl * line_h(size) + 0.16 * k))
    return cw, hs


def blk_h(b: Blk, k):
    w = W_IN - 2 * MX
    if b.kind == "msg":
        return n_lines(b.text, SIZE["msg"] * k, w - 0.4) * line_h(SIZE["msg"] * k) + 0.36 * k
    if b.kind == "table":
        return sum(table_geom(b.rows, w, k)[1])
    if b.kind == "hr":
        return 0.02
    size = SIZE["tail" if b.kind == "tail" else b.kind] * k
    ww = w - INDENT.get(b.kind, 0.0) - 0.35 - (text_w(b.right, size) + 0.2 if b.right else 0)
    return n_lines(b.text, size, ww) * line_h(size)


GAP_BEFORE = {"msg": 0.0, "l1": 0.2, "sub": 0.06, "foot": 0.02, "table": 0.1, "hr": 0.16, "tail": 0.14}


def total_h(blks, k):
    h = 0.0
    for i, b in enumerate(blks):
        h += (GAP_BEFORE[b.kind] * k if i else 0) + blk_h(b, k)
    # 표 바로 다음 블록은 조금 더 띄운다
    return h


def fit(blks):
    avail = BOTTOM - TOP
    k = 1.1
    while k > 0.6 and total_h(blks, k) > avail:
        k -= 0.02
    return k, avail - total_h(blks, k)


# ── 그리기 ────────────────────────────────────────────────
def runs(p, text, size, color=BLACK, bold=False):
    for part in re.split(r"(\*\*[^*]+\*\*|__[^_]+__)", text):
        if not part:
            continue
        c, b = color, bold
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            part, c, b = part[2:-2], BLUE, True
        elif part.startswith("__") and part.endswith("__") and len(part) > 4:
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


def tbox(s, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.03)
    tf.margin_top = tf.margin_bottom = 0
    return tf


def _nostyle(shp):
    st = shp._element.find(qn("p:style"))
    if st is not None:
        shp._element.remove(st)


def label_line(p, text, size, color=BLACK):
    """'라벨: 값' → 라벨 굵게. 값의 **/__ 표기는 그대로 살린다."""
    m = re.match(r"^([^:：]{1,14})([:：])(.*)$", text)
    if m and "**" not in m.group(1):
        runs(p, m.group(1).strip() + ":", size, color, True)
        runs(p, " " + m.group(3).strip(), size, color)
    else:
        runs(p, text, size, color)


def cell_border(cell, color="BFBFBF", w=Pt(0.75)):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for el in tcPr.findall(qn(tag)):
            tcPr.remove(el)
        ln = etree.SubElement(tcPr, qn(tag), w=str(int(w)))
        sf = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr"), val=color)
    for t in ("a:solidFill", "a:noFill"):
        for el in tcPr.findall(qn(t)):
            tcPr.remove(el)
            tcPr.append(el)


def draw_table(s, x, y, w, rows, k):
    cw, hs = table_geom(rows, w, k)
    nc = len(cw)
    rows = [r + [""] * (nc - len(r)) for r in rows]
    gt = s.shapes.add_table(len(rows), nc, Inches(x), Inches(y), Inches(w), Inches(sum(hs)))
    tbl = gt.table
    tblPr = tbl._tbl.tblPr
    for a in ("firstRow", "bandRow"):
        tblPr.set(a, "0")
    for j, c in enumerate(cw):
        tbl.columns[j].width = Inches(c)
    size = SIZE["table"] * k
    for i, row in enumerate(rows):
        tbl.rows[i].height = Inches(hs[i])
        for j, val in enumerate(row):
            c = tbl.cell(i, j)
            c.margin_left = c.margin_right = Inches(0.04)
            c.margin_top = c.margin_bottom = 0
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = c.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            runs(p, val, size)
            c.fill.solid()
            if i == 0:
                c.fill.fore_color.rgb = RGBColor(0xCF, 0xE5, 0xFF)
            elif j == 0 or val.strip():
                c.fill.fore_color.rgb = RGBColor(0xF2, 0xF2, 0xF2)
            else:
                c.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            cell_border(c)
    return sum(hs)


def build(spec: Path, out: Path, template: Path | None):
    meta, blks = parse(spec)
    if not blks:
        sys.exit("문안에서 내용을 찾지 못했습니다.")
    if template and template.exists():
        prs = Presentation(str(template))
        layout = prs.slide_layouts[0]
        for l in prs.slide_layouts:
            if len(l.placeholders) <= 1:
                layout = l
                break
        lst = prs.slides._sldIdLst
        for sld in list(lst):
            prs.part.drop_rel(sld.get(qn("r:id")))
            lst.remove(sld)
    else:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(W_IN), Inches(H_IN)
        layout = prs.slide_layouts[6]
    s = prs.slides.add_slide(layout)
    for ph in list(s.placeholders):
        ph._element.getparent().remove(ph._element)

    k, spare = fit(blks)
    over = total_h(blks, k) > BOTTOM - TOP
    # 남는 높이는 ● 사이 간격에 나눠 아래까지 고르게 채운다(상한 있음)
    n_l1 = sum(1 for b in blks if b.kind in ("l1", "table")) or 1
    extra = 0.0 if over else min(0.18, spare * 0.7 / n_l1)

    w = W_IN - 2 * MX
    y = TOP
    for i, b in enumerate(blks):
        if i:
            y += GAP_BEFORE[b.kind] * k + (extra if b.kind in ("l1", "table") else 0)
        h = blk_h(b, k)
        if b.kind == "msg":
            box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(MX), Inches(y), Inches(w), Inches(h))
            _nostyle(box)
            box.fill.background()
            box.line.color.rgb = RGBColor(0xBF, 0xBF, 0xBF)
            box.line.width = Pt(3)
            tf = box.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            runs(p, b.text, SIZE["msg"] * k, NAVY, False)
        elif b.kind == "table":
            draw_table(s, MX, y, w, b.rows, k)
        elif b.kind == "hr":
            ln = s.shapes.add_connector(1, Inches(MX), Inches(y), Inches(MX + w), Inches(y))
            _nostyle(ln)
            ln.line.color.rgb = RGBColor(0xBF, 0xBF, 0xBF)
            ln.line.width = Pt(0.75)
        else:
            size = SIZE["tail" if b.kind == "tail" else b.kind] * k
            ind = INDENT.get(b.kind, 0.0)
            rw = text_w(b.right, size) + 0.2 if b.right else 0.0
            if b.kind in ("l1", "tail"):
                mark = tbox(s, MX, y, 0.35, line_h(size))
                runs(mark.paragraphs[0], "●" if b.kind == "l1" else "■", size)
                tx = MX + 0.35
            else:
                tx = MX + ind
            tf = tbox(s, tx, y, w - (tx - MX) - rw, h)
            p = tf.paragraphs[0]
            p.line_spacing = 1.0
            if b.kind == "foot":
                runs(p, b.text, size, GRAY)
            elif b.kind == "sub" and b.text.startswith("▶"):
                runs(p, "▶ ", size)
                label_line(p, b.text[1:].strip(), size)
            elif b.kind == "sub":
                runs(p, b.text, size)
            else:
                label_line(p, b.text, size)
            if b.right:
                rt = tbox(s, MX + w - rw, y, rw, line_h(size))
                rp = rt.paragraphs[0]
                rp.alignment = PP_ALIGN.RIGHT
                runs(rp, b.right, size * 0.85, BLACK)
        y += h
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    print(f"생성: {out}  (글자 배율 {k:.2f})")
    if over:
        print("  ⚠ 한 장에 다 안 들어갈 수 있음 → 각주·하위 줄을 줄이거나 표 행을 합치세요")


def main():
    ap = argparse.ArgumentParser(description="안내문 문안 → 텍스트 중심 1-page 장표")
    ap.add_argument("spec")
    ap.add_argument("-o", "--out")
    ap.add_argument("--template", default=None, help="회사 양식 pptx(선택). 기본: templates/row.pptx가 있으면 사용")
    a = ap.parse_args()
    spec = Path(a.spec)
    meta, _ = _front(spec.read_text(encoding="utf-8"))
    out = Path(a.out) if a.out else spec.with_name(re.sub(r'[\\/:*?"<>|]', "", meta.get("title", spec.stem)) + ".pptx")
    tpl = Path(a.template) if a.template else (TEMPLATE if TEMPLATE.exists() else None)
    build(spec, out, tpl)


if __name__ == "__main__":
    main()
