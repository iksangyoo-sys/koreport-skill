"""문장 패턴 데이터 하나로 스킬 참고 파일과 학습자용 웹 도구를 함께 만든다.

    python scripts/build_patterns.py

입력:  assets/sentence-patterns.json        (원천 데이터 — 패턴을 고칠 때는 이 파일만 고친다)
       assets/sentence-patterns-guide.md    (학습자 사용안내서 — 웹 도구의 '사용안내서' 창에 들어간다)
       scripts/sentence-patterns.template.html
출력:  references/13-sentence-patterns.md  (에이전트가 읽는 참고 파일)
       assets/sentence-patterns.html       (학습자에게 주는 '문장 공식 찾기', 인터넷 없이 동작)
"""
import html
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "assets" / "sentence-patterns.json"
GUIDE = ROOT / "assets" / "sentence-patterns-guide.md"
TEMPLATE = ROOT / "scripts" / "sentence-patterns.template.html"
OUT_MD = ROOT / "references" / "13-sentence-patterns.md"
OUT_HTML = ROOT / "assets" / "sentence-patterns.html"

# 받침 유무를 모를 때(빈칸) 보여 줄 조사 표기
JOSA_PLACEHOLDER = {
    "은": "은(는)", "이": "이(가)", "을": "을(를)", "과": "과(와)", "으로": "(으)로",
    "이란": "(이)란", "이나": "(이)나", "이었으나": "(이)었으나",
}
SLOT = re.compile(r"\{([^}:]+)(?::([^}]+))?\}")


def render_placeholder(template):
    """'{대상:은} {행동}' → '(대상)은(는) (행동)'"""
    return SLOT.sub(lambda m: f"({m.group(1)})" + (JOSA_PLACEHOLDER[m.group(2)] if m.group(2) else ""), template)


def build_md(d):
    rel = {r["id"]: r for r in d["relations"]}
    areas = [a["id"] for a in d["areas"]]
    by_rel = {}
    for p in d["patterns"]:
        by_rel.setdefault(p["relation"], []).append(p)

    out = []
    w = out.append
    w("# 13. 문장 패턴 공식 — 보고 단계와 문장 관계로 공식을 골라 빈칸을 채운다")
    w("")
    w("> **언제 읽나:** 헤드라인·본문 문장을 조립할 때 어떤 문형을 쓸지 막힐 때, AI 초안이 번역투·서술형이라 한국 기업 보고서 문체로 바꿔야 할 때, 문장 진단에서 빠진 요소를 패턴 이름으로 짚을 때 연다. 키워드 조립·간결화는 09편, 종결·표기 교정은 10편을 함께 본다.")
    w("")
    w(f"> **자료:** {d['about']} 원천 데이터는 `assets/sentence-patterns.json`이고, 이 파일은 `scripts/build_patterns.py`로 만든다. 직접 고치지 말고 JSON을 고친 뒤 다시 만든다.")
    w("")
    w("## 핵심 원칙")
    w("")
    w("1. **메시지가 먼저, 공식은 나중이다.** 독자에게 무엇을 말할지(키워드·메시지)가 정해진 뒤에 공식을 고른다. 공식부터 고르면 형식만 그럴듯하고 내용이 빈 문장이 나온다.")
    w("2. **두 가지만 정하면 공식이 나온다.** ① 이 문장이 보고서의 어느 단계(배경·현상·개선방안·실행계획)에 들어가는가 ② 무엇과 무엇의 관계(정의·비교·전환·인과…)를 말하는가.")
    w("3. **기본형은 킬러문장이다.** `(대상)으로 (상황)에서 (목적)을 위해 (수단)을 통해 (행동)` 다섯 요소 가운데 필요한 것만 골라 조합한다. 기업 보고서 문장의 절반 이상이 이 조합으로 쓰인다.")
    w("4. **빈칸은 명사와 수치로 채우고, 끝은 명사형으로 맺는다.** '구축, 운영, 강화, 필요, 추진'처럼 행동 명사로 끝내고, 한 문서 안에서 종결 패턴을 하나로 맞춘다(10편).")
    w("5. **예문은 구조만 빌린다.** 예문의 내용·수치를 근거로 쓰지 않는다. 수치는 O로 가려 두었다.")
    w("")
    w("## 방법과 도구")
    w("")
    w("### 1. 공식 고르기 — 보고 단계 × 문장 관계")
    w("")
    w("| 관계 \\ 단계 | " + " | ".join(areas) + " |")
    w("|---|" + "---|" * len(areas))
    for r in d["relations"]:
        cells = []
        for a in areas:
            ids = [p["id"] for p in by_rel.get(r["id"], []) if a in p["areas"]]
            cells.append(", ".join(ids) if ids else "")
        w(f"| {r['name']} | " + " | ".join(cells) + " |")
    w("")
    w("단계별 뜻: " + " / ".join(f"**{a['id']}** {a['desc']}" for a in d["areas"]))
    w("")
    w("### 2. 문장 관계 11가지")
    w("")
    w("| 관계 | 공식 | 쓰임 | 이런 말을 하고 싶을 때 |")
    w("|---|---|---|---|")
    for r in d["relations"]:
        w(f"| {r['name']} | {r['formula']} | {r['desc']} | {', '.join(r['keywords'][:6])} |")
    w("")
    w("### 3. 빈칸(요소) 뜻")
    w("")
    w("| 빈칸 | 채울 내용 |")
    w("|---|---|")
    for k, v in d["slots"].items():
        w(f"| {k} | {v} |")
    w("")
    w(f"조사: {d['josa_note']} 아래 공식의 `은(는)`, `(으)로` 같은 표기가 그 자리다.")
    w("")
    w("### 4. 패턴 22")
    w("")
    for r in d["relations"]:
        ps = by_rel.get(r["id"], [])
        if not ps:
            continue
        w(f"#### {r['name']} — {r['formula']}")
        w("")
        if r.get("note"):
            w(f"{r['desc']} {r['note']}")
            w("")
        for p in ps:
            w(f"**{p['id']}. {p['combo']}** · 단계: {', '.join(p['areas'])}")
            w("")
            w(f"- 공식: `{render_placeholder(p['template'])}`")
            if p["variants"]:
                w("- 변형: " + " / ".join(f"`{render_placeholder(v)}`" for v in p["variants"]))
            w(f"- 요령: {p['tip']}")
            w("- 예문:")
            for e in p["examples"]:
                w(f"  - {e}")
            w("")
    w("### 5. 에이전트 사용법")
    w("")
    w("- **쓸 때:** 문장마다 단계와 관계를 먼저 정하고 1번 표에서 패턴을 고른 뒤, 빈칸을 자료의 명사·수치로 채운다. 자료에 없는 빈칸은 지어내지 말고 \"확인 필요\"로 둔다.")
    w("- **진단할 때:** 문장이 어느 패턴인지 이름을 붙이고, 빠진 요소를 짚는다. 예: 인과인데 결과가 없다, 전환인데 이전 상태가 없다, 비교인데 상반되는 짝이 없다, 개선인데 '검토 필요'로 끝나 행동이 없다.")
    w("- **AI 초안을 고칠 때:** '~하기 때문에 ~할 필요가 있습니다' 같은 서술형·번역투를 관계에 맞는 패턴으로 바꾸고 명사형으로 맺는다(전후 사례 참고).")
    w("- **학습자에게 도구를 줄 때:** `assets/sentence-patterns.html`을 건넨다. 파일 하나로 인터넷 없이 열리며, 하고 싶은 말(관계)을 고르고 문장 속 빈칸을 채우면 조사까지 맞춰 문장이 조립된다.")
    w("")
    w("## 전후 사례")
    w("")
    w("前은 AI 초안에 흔한 서술형·번역투 문장을 예시로 쓴 것이고, 後는 위 예문이다.")
    w("")
    cases = [
        ("A. 서술형 → 목적+대상+행동 (K4)",
         "글로벌 경영 환경이 빠르게 변화하고 있기 때문에, 우리는 분산되어 있는 정보들을 하나로 통합하는 것을 통해 글로벌 경영정보 체계를 구축할 필요가 있습니다.",
         "급변하는 글로벌 경영환경에 대응하기 위해 분산된 정보를 하나로 통합하여 글로벌 경영정보 체계 구축",
         "'~때문에 ~할 필요가 있습니다'를 목적(대응하기 위해) + 대상(분산된 정보) + 행동(체계 구축)으로 조립하고 명사형으로 맺었다."),
        ("B. 두 문장 → 비교 (C1)",
         "제안 제도는 원래 아이디어 실행에 대해 보상하기 위한 제도였습니다. 그런데 현재는 성과보다는 참여에 대해 보상하는 방향으로 바뀌었습니다.",
         "제안은 아이디어 실행에 대한 보상 제도이나, 당초 취지인 ‘성과’ 보다는 ‘참여’ 보상으로 퇴색",
         "상황(취지)과 이슈(퇴색)를 '~이나'로 맞세워 한 문장에 문제를 드러냈다. 상반어 짝: 성과/참여."),
        ("C. 당위 나열 → 전환 (T3)",
         "우리는 부문별·계열사별로 최적화하는 것에서 벗어나 글로벌 차원의 최적화를 추구해야 하며, 고객에게는 가격보다는 가치로 경쟁력을 강화해야 합니다.",
         "부문·계열사별 최적화가 아닌 「글로벌 최적화」, 고객에 대해 가격이 아닌 「가치 경쟁력」 강화",
         "'~에서 벗어나 ~해야 한다'를 'A가 아닌 B'로 바꿔 버리는 쪽과 가는 쪽을 대비했다. 핵심어는 「 」로 강조했다."),
    ]
    for title, before, after, why in cases:
        w(f"### {title}")
        w("")
        w(f"- 前: {before}")
        w(f"- 後: {after}")
        w(f"- 왜: {why}")
        w("")
    w("## 점검 질문")
    w("")
    for q in [
        "공식을 고르기 전에 이 문장의 메시지(독자에게 할 말)가 정해져 있는가?",
        "이 문장의 단계(배경·현상·개선방안·실행계획)와 관계(정의·비교·전환·인과…)를 말할 수 있는가?",
        "고른 패턴의 빈칸이 모두 채워졌는가? (인과의 결과, 전환의 이전 상태, 비교의 상반 짝, 개선의 구체 행동)",
        "빈칸을 막연한 말이 아니라 명사·수치로 채웠는가?",
        "문장이 명사형으로 끝나고, 같은 문서 안에서 종결 패턴이 하나로 맞춰져 있는가?",
        "'통해', '위해'가 줄마다 반복되지 않는가? 반복되면 변형 공식이나 수단·목적 자리 바꾸기를 쓴다(09편).",
        "예문의 내용이나 수치를 실제 근거처럼 가져다 쓰지 않았는가?",
    ]:
        w(f"- [ ] {q}")
    w("")
    return "\n".join(out)


def _inline(text):
    """**굵게**, `코드`만 처리한다 (안내서에 쓰는 것만)."""
    out = html.escape(text, quote=False)
    out = re.sub(r"`([^`]+)`", r"<code>\1</code>", out)
    return re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", out)


def guide_to_html(md):
    """안내서용 최소 마크다운 변환: 제목(#, ##), 문단, 목록(-, 1., 4칸 들여쓴 하위 목록), 표, ``` 코드 블록."""
    lines = md.splitlines()
    out, i = [], 0
    while i < len(lines):
        line = lines[i]
        if not line.strip():
            i += 1
        elif line.startswith("```"):
            i += 1
            code = []
            while i < len(lines) and not lines[i].startswith("```"):
                code.append(lines[i]); i += 1
            i += 1
            out.append('<div class="code"><button type="button" class="btn ghost copy-code">복사</button><pre>'
                       + html.escape("\n".join(code), quote=False) + "</pre></div>")
        elif line.startswith("## "):
            out.append(f"<h3>{_inline(line[3:])}</h3>"); i += 1
        elif line.startswith("# "):
            i += 1  # 문서 제목은 창 머리글이 대신한다
        elif line.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].startswith("|"):
                rows.append([c.strip() for c in lines[i].strip().strip("|").split("|")]); i += 1
            head, body = rows[0], [r for r in rows[2:]]
            out.append('<div class="table-wrap"><table><thead><tr>' + "".join(f"<th>{_inline(c)}</th>" for c in head)
                       + "</tr></thead><tbody>" + "".join("<tr>" + "".join(f"<td>{_inline(c)}</td>" for c in r) + "</tr>" for r in body)
                       + "</tbody></table></div>")
        elif re.match(r"(- |\d+\. )", line):
            tag = "ol" if line[0].isdigit() else "ul"
            items = []
            while i < len(lines) and (re.match(r"(- |\d+\. )", lines[i]) or lines[i].startswith("    ")):
                if lines[i].startswith("    "):
                    items[-1][1].append(re.sub(r"^\s+(- |\d+\. )", "", lines[i]))
                else:
                    items.append([re.sub(r"^(- |\d+\. )", "", lines[i]), []])
                i += 1
            out.append(f"<{tag}>" + "".join(
                f"<li>{_inline(t)}" + ("<ul>" + "".join(f"<li>{_inline(c)}</li>" for c in sub) + "</ul>" if sub else "") + "</li>"
                for t, sub in items) + f"</{tag}>")
        else:
            para = []
            while i < len(lines) and lines[i].strip() and not re.match(r"(#|\||```|- |\d+\. )", lines[i]):
                para.append(lines[i].strip()); i += 1
            out.append(f"<p>{_inline(' '.join(para))}</p>")
    return "\n".join(out)


def build_html(d):
    tpl = TEMPLATE.read_text(encoding="utf-8")
    payload = json.dumps(d, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    marker = "/*__PATTERN_DATA__*/null"
    guide_marker = "<!--__GUIDE_HTML__-->"
    assert marker in tpl, "템플릿에 데이터 자리표시가 없습니다"
    assert guide_marker in tpl, "템플릿에 안내서 자리표시가 없습니다"
    guide = guide_to_html(GUIDE.read_text(encoding="utf-8"))
    return tpl.replace(marker, payload).replace(guide_marker, guide)


def main():
    d = json.loads(DATA.read_text(encoding="utf-8"))
    assert len(d["patterns"]) == 22, f"패턴 수가 22개가 아닙니다: {len(d['patterns'])}"
    rel_ids = {r["id"] for r in d["relations"]}
    for p in d["patterns"]:
        assert p["relation"] in rel_ids, p["id"]
        for t in [p["template"], *p["variants"]]:
            for m in SLOT.finditer(t):
                assert m.group(1) in p["slots"], f"{p['id']}: 빈칸 '{m.group(1)}'이 slots에 없습니다"
                assert not m.group(2) or m.group(2) in JOSA_PLACEHOLDER, f"{p['id']}: 알 수 없는 조사 '{m.group(2)}'"
    OUT_MD.write_text(build_md(d), encoding="utf-8")
    OUT_HTML.write_text(build_html(d), encoding="utf-8")
    print(f"만들었습니다: {OUT_MD.relative_to(ROOT)}, {OUT_HTML.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
