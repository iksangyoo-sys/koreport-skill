#!/usr/bin/env python3
"""
koreport-skill — 기본 장표(spec_to_pptx.py 결과)의 일부 장을 PPT Master로 만든 장표로 바꿔 끼운다.

PPT Master는 시각화가 꼭 필요한 장(@ppt-master 표시)만 따로 만들고,
나머지는 저비용 렌더러 결과를 그대로 쓰기 위한 도구다.

사용법:
    python merge_slides.py <base.pptx> <visual.pptx> --map 4=1,6=2 [-o final.pptx]

    --map  "기본 덱의 쪽=시각화 덱의 쪽" (1부터, 표지 포함 실제 슬라이드 순서)
           기본 덱의 해당 장을 시각화 덱의 장으로 교체한다.

제약:
    - PPT Master의 flat 출력(모든 도형이 슬라이드 안에 있음)을 전제로 한다.
      structured(Master/Layout 상속) 출력이면 상속 서식이 빠질 수 있다.
    - 이미지·차트·하이퍼링크·노트는 함께 옮긴다. 애니메이션·전환은 원본 XML을 그대로 옮기며
      오디오·비디오 등 미디어는 옮기지만 재생 설정은 확인이 필요하다.
의존성: python-pptx
"""

from __future__ import annotations

import argparse
import copy
import re
import sys
from io import BytesIO
from pathlib import Path

from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SKIP_RELS = {RT.SLIDE_LAYOUT, RT.NOTES_SLIDE}


def _rename_tree(part, dest_pkg, seen: set):
    """복사해 올 파트(차트·임베드 등)의 이름이 기본 덱과 겹치지 않게 다시 붙인다."""
    if id(part) in seen:
        return
    seen.add(id(part))
    tmpl = re.sub(r"\d+(?=\.[^.]+$)", "%d", str(part.partname))
    if "%d" not in tmpl:
        stem, ext = str(part.partname).rsplit(".", 1)
        tmpl = f"{stem}%d.{ext}"
    part.partname = dest_pkg.next_partname(tmpl)
    for rel in part.rels.values():
        if not rel.is_external:
            _rename_tree(rel.target_part, dest_pkg, seen)


def copy_slide(dest_prs, src_slide):
    layout = dest_prs.slide_layouts[6] if len(dest_prs.slide_layouts) > 6 else dest_prs.slide_layouts[-1]
    new = dest_prs.slides.add_slide(layout)
    # 레이아웃 자리표시자 제거
    for shp in list(new.shapes):
        shp._element.getparent().remove(shp._element)

    dest_part, src_part = new.part, src_slide.part
    dest_pkg = dest_prs.part.package
    rid_map: dict[str, str] = {}
    seen: set = set()
    for rid, rel in src_part.rels.items():
        if rel.reltype in SKIP_RELS:
            continue
        if rel.is_external:
            rid_map[rid] = dest_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        elif rel.reltype == RT.IMAGE:
            _, new_rid = dest_part.get_or_add_image_part(BytesIO(rel.target_part.blob))
            rid_map[rid] = new_rid
        else:
            _rename_tree(rel.target_part, dest_pkg, seen)
            rid_map[rid] = dest_part.relate_to(rel.target_part, rel.reltype)

    src_csld = src_slide._element.find(qn("p:cSld"))
    dst_csld = new._element.find(qn("p:cSld"))
    # 배경
    src_bg = src_csld.find(qn("p:bg"))
    if src_bg is not None:
        dst_csld.insert(0, copy.deepcopy(src_bg))
    # 도형 트리
    dst_tree = new.shapes._spTree
    for el in src_slide.shapes._spTree:
        if el.tag in (qn("p:nvGrpSpPr"), qn("p:grpSpPr")):
            continue
        dst_tree.append(copy.deepcopy(el))
    # 전환·애니메이션
    for tag in ("p:transition", "p:timing"):
        el = src_slide._element.find(qn(tag))
        if el is not None:
            new._element.append(copy.deepcopy(el))
    # r:id 재매핑
    for el in new._element.iter():
        for attr, val in list(el.attrib.items()):
            if attr.startswith("{%s}" % R_NS) and val in rid_map:
                el.set(attr, rid_map[val])
    # 노트
    if src_slide.has_notes_slide:
        txt = src_slide.notes_slide.notes_text_frame.text
        if txt:
            new.notes_slide.notes_text_frame.text = txt
    return new


def move_slide(prs, old_index: int, new_index: int):
    lst = prs.slides._sldIdLst
    items = list(lst)
    el = items[old_index]
    lst.remove(el)
    lst.insert(new_index, el)


def delete_slide(prs, index: int):
    lst = prs.slides._sldIdLst
    el = list(lst)[index]
    prs.part.drop_rel(el.get(qn("r:id")))
    lst.remove(el)


def main():
    ap = argparse.ArgumentParser(description="기본 덱의 일부 장을 PPT Master 장표로 교체")
    ap.add_argument("base", type=Path)
    ap.add_argument("visual", type=Path)
    ap.add_argument("--map", required=True, help="예: 4=1,6=2 (기본 쪽=시각화 쪽)")
    ap.add_argument("-o", "--out", type=Path)
    args = ap.parse_args()

    pairs = []
    for item in args.map.split(","):
        b, v = item.split("=")
        pairs.append((int(b), int(v)))

    base = Presentation(args.base)
    vis = Presentation(args.visual)
    if (base.slide_width, base.slide_height) != (vis.slide_width, vis.slide_height):
        print(f"경고: 슬라이드 크기가 다릅니다 (기본 {base.slide_width}x{base.slide_height}, "
              f"시각화 {vis.slide_width}x{vis.slide_height}). 도형 위치가 어긋날 수 있습니다.")
    n_base, n_vis = len(base.slides), len(vis.slides)
    for b, v in pairs:
        if not (1 <= b <= n_base and 1 <= v <= n_vis):
            sys.exit(f"--map 범위 오류: {b}={v} (기본 {n_base}장, 시각화 {n_vis}장)")

    # 뒤쪽부터 교체해야 앞쪽 번호가 흔들리지 않는다.
    for b, v in sorted(pairs, reverse=True):
        copy_slide(base, vis.slides[v - 1])
        move_slide(base, len(base.slides) - 1, b - 1)
        delete_slide(base, b)  # 밀려난 원래 장
    out = args.out or args.base.with_name(args.base.stem + "_merged.pptx")
    base.save(out)
    print(f"생성: {out}  (교체 {len(pairs)}장)")


if __name__ == "__main__":
    main()
