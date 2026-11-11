# koreport-skill

[![skills.sh](https://skills.sh/b/iksangyoo-sys/koreport-skill)](https://skills.sh/iksangyoo-sys/koreport-skill)

한국어 업무 보고서를 **koreport 방법론으로 문안부터 확정**하고, PPT는 **필요한 만큼만** 만드는 결합 스킬입니다.

- 문안(독자 니즈 → 메시지 → 목차·헤드라인 → 본문)은 `report_spec.md` 한 파일로 확정합니다.
- 기본 장표는 회사 양식(`templates/row.pptx` 가로, `col.pptx` 세로, `bullet.pptx` 도형 견본)으로 만듭니다. 사용자가 요청하지 않아도 적용합니다. **회사 양식 파일은 이 저장소에 포함하지 않으니** 각자 회사 양식을 같은 이름으로 `templates/`에 넣어 쓰세요(`templates/README.md` 참고). `scripts/spec_to_onepage.py`(가로)·`spec_to_company.py`(세로)가 문안을 그대로 렌더링하므로 AI 없이 PC에서 돌려도 됩니다(비용 0원).
- 고정 블루 양식은 사용자가 요청할 때만 씁니다. 필요하면 `scripts/spec_to_pptx.py`가 한국 기업 1-page 양식(블루 단색, 맑은 고딕, 장 4층 구조, 기호 위계)으로 바로 렌더링합니다. LLM이 장표를 그리지 않으므로 비용이 거의 들지 않습니다.
- 개념도·프로세스·차트처럼 시각화가 메시지를 결정하는 장만 `@ppt-master`로 표시해 PPT Master로 그리고, `scripts/merge_slides.py`로 기본 덱에 끼워 넣습니다.

## 빠른 사용

"회사 양식에 맞춰 아래 내용을 작성해 줘"라고 하면 스킬이 먼저 **가로(row)/세로(col)**와 **장 수**를 묻고, 답에 맞춰 `templates`의 견본 배치대로 바로 만듭니다.

```bash
# 1) 가로 = row.pptx 견본 배치 (묶음 문장·회색선 박스·라벨 바·표·원형 라벨·하단 따옴표)
python scripts/spec_to_onepage.py templates/onepage_spec_template.md -o out_row.pptx --pages 1
# 2) 세로 = col.pptx 견본 배치 (A4, 섹션 + 둥근 라벨 행)
python scripts/spec_to_company.py report_spec.md -o out_col.pptx --layout col
# 3) 문안 점검 / @ppt-master 장 브리프 추출 (T2)
python scripts/spec_to_pptx.py report_spec.md --check
python scripts/spec_to_pptx.py report_spec.md --extract-visual visual_brief.md
# 4) PPT Master 결과를 기본 덱에 병합
python scripts/merge_slides.py out.pptx ppt_master_result.pptx --map 4=1 -o final.pptx
```

필요 패키지: `python-pptx` (T1·병합). PPT Master를 쓰는 T2·T3는 [hugohe3/ppt-master](https://github.com/hugohe3/ppt-master)를 따로 설치해 `ppt-master/` 폴더로 두세요.

## 구성

| 경로 | 내용 |
|---|---|
| `SKILL.md` | 전체 흐름과 출력 단계(T0~T3) |
| `workflows/01-report-spec.md` | A단계: koreport로 문안 확정 |
| `workflows/02-spec-format.md` | `report_spec.md` 문법 |
| `workflows/03-pptx-output.md` | B단계: 장표 생성, PPT Master 연결, 수정 흐름 |
| `templates/onepage_spec_template.md` | 가로(row) 견본 배치 빈 문안 |
| `templates/report_spec_template.md` | 여러 장·세로(col) 빈 문안 |
| `templates/README.md` | 회사 양식 파일(row·col·bullet) 넣는 법 |
| `scripts/spec_to_onepage.py` | 가로 회사 양식(row.pptx 견본 배치) 렌더러 — 가로 기본 |
| `scripts/spec_to_company.py` | 세로 회사 양식(col.pptx) 렌더러 — 세로 기본, row 표지+여러 장 |
| `scripts/spec_to_pptx.py` | 블루 양식 렌더러 · 양식 점검기 · 브리프 추출 |
| `scripts/merge_slides.py` | 장표 교체 병합 |
| `koreport/` | koreport 방법론 12편 + 전용 스킬 9개 (하위 스킬 파일명은 `GUIDE.md`) |
| `index.html`, `assets/sentence-patterns.html` | 웹 도구 「보고서 문장 공식」 (https://iksangyoo-sys.github.io/koreport-skill/) |
| `koreport/scripts/build_patterns.py` | 문장 패턴·어휘 JSON을 고친 뒤 웹 도구와 두 전용 스킬(GUIDE.md)을 다시 만드는 스크립트 |

## 원본과 라이선스

- koreport 방법론(「AI가 못 쓰는 보고서 쓰기」 교안)은 `koreport/`에 들어 있습니다.
- PPT Master(T2·T3)는 이 저장소에 포함하지 않습니다. 쓰려면 [hugohe3/ppt-master](https://github.com/hugohe3/ppt-master) v6.9.0(MIT, Copyright (c) 2025-2026 Hugo He)의 `skills/ppt-master`를 `ppt-master/` 폴더로 넣으세요. 내부 `attribution_guard.py`가 `SKILL.md`·`LICENSE` 등을 검사하므로 그 폴더의 파일은 수정하지 마세요. 결합에 필요한 규칙은 모두 바깥(`SKILL.md`, `workflows/`)에 두었습니다.

## 설치 참고

`npx skills add iksangyoo-sys/koreport-skill`로 설치하거나, 이 폴더째 스킬 폴더(예: `~/.claude/skills/koreport-skill`)에 두면 됩니다. claude.ai에 zip으로 올릴 때는 ZIP 안 파일명을 영문으로 두고 `SKILL.md`는 맨 위 하나만 둡니다(그래서 `koreport/` 하위 스킬 파일은 `GUIDE.md`로 바꿔 두었습니다).
