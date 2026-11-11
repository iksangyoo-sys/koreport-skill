---
name: koreport-skill
description: 한국어 업무 보고서를 koreport 방법론(독자 니즈 → 메시지 → 스토리라인 → 헤드라인·묶고–풀기 → 1-page 장표 양식)으로 먼저 기획·문안 확정한 뒤, 필요한 만큼만 PPTX로 만든다. 문안은 report_spec.md 한 파일로 확정하고, 기본 장표는 스크립트가 1-page 양식으로 바로 렌더링하며, 시각화가 꼭 필요한 장만 PPT Master로 그려 끼운다. 보고서를 PPT·장표·슬라이드로 만들어 달라, 보고서 기획부터 PPT까지, 우리 회사 장표 양식으로 PPT, 저비용·빠르게 보고서 PPT를 요청받을 때 사용한다. 문안만 고치는 일은 koreport, 디자인 중심 발표 자료는 PPT Master를 단독으로 쓴다.
---

# koreport-skill — 문안은 koreport로, 장표는 필요한 만큼만

보고서 PPT의 품질은 장표 디자인이 아니라 **메시지와 문안**에서 나오고, 비용은 장표를 LLM이 한 장씩 그릴 때 가장 크게 든다. 이 스킬은 두 일을 떼어 놓는다.

1. **문안 확정 (koreport):** 독자·메시지·목차·헤드라인·본문을 `report_spec.md` 한 파일로 확정한다. 토큰은 대부분 여기에 쓴다.
2. **장표 생성 (단계별 선택):** 확정 문안을 스크립트가 1-page 양식으로 렌더링한다(LLM 토큰 거의 0). 시각화가 꼭 필요한 장만 `@ppt-master`로 표시해 PPT Master로 그리고 병합한다.

```
koreport-skill/   (claude.ai 업로드용 경량판: ppt-master 사본 제외)
├── SKILL.md                     ← 지금 이 파일 (전체 흐름)
├── workflows/
│   ├── 01-report-spec.md        ← A단계: koreport로 문안 확정하는 절차
│   ├── 02-spec-format.md        ← report_spec.md 문법
│   └── 03-pptx-output.md        ← B단계: 출력 단계(T0~T3)와 PPT Master 연결
├── templates/
│   ├── report_spec_template.md
│   ├── col.pptx                 ← 회사 세로(A4) 1장 양식 (세로양식)
│   ├── row.pptx                 ← 회사 가로 1장 양식, 예시 70장 (가로양식)
│   └── bullet.pptx              ← 가로 도형·레이아웃 견본 16장 (도형양식)
├── examples/
│   ├── sample_report.md (+ .pptx)
│   ├── onepage_vertical_meeting/   ← 세로양식 1장 생성 예시 (build_onepage.py + 문안)
│   └── onepage_landscape_privacy/  ← 가로양식 1장 생성 예시 (build_onepage.py + 문안)
├── scripts/
│   ├── spec_to_company.py       ← 문안 → 회사 양식(row 가로 / col A4 세로) PPTX  ★기본 렌더러
│   ├── spec_to_pptx.py          ← 문안 → 고정 블루 양식 PPTX, 양식 점검, PPT Master 브리프 추출
│   └── merge_slides.py          ← PPT Master 장표를 기본 덱에 교체 삽입
└── koreport/                    ← koreport-skill 원본 사본 (방법론·전용 스킬)
```

## 이 경량판에서 쓸 수 있는 것

- **T0(문안)·T1(스크립트 장표)만 동작한다.** `ppt-master/` 사본은 용량 때문에 뺐으므로 T2·T3, `source_to_md.py`, Edit Native PPTX 경로는 쓸 수 없다. 사용자가 요청하면 PC의 Claude Code 등 ppt-master가 있는 환경에서 하도록 한 줄로 안내한다. 원자료(PDF·DOCX·PPTX)는 직접 읽는다.
- **장표 기본 양식 = 회사 양식 (사용자가 따로 말하지 않아도 적용):** 장표를 만들 때는 항상 회사 양식(templates의 row·col·bullet)을 쓴다. 가로 보고서는 `row.pptx`(기본), A4 세로 문서(회의 결과·간담회 결과 등 한 장짜리 세로 보고, 또는 사용자가 세로를 요청할 때)는 `col.pptx`. 구조도·흐름도·일정표가 필요한 장은 먼저 `bullet.pptx` 견본을 쓴다. `scripts/spec_to_pptx.py`의 고정 블루 양식은 사용자가 "기본 양식"·"블루 양식"을 명시적으로 요청할 때만 렌더링에 쓰고, 평소에는 `--check`(문안 점검)와 `--extract-visual`에만 쓴다.
- **회사 양식으로 만드는 방법 (기본):** 문안을 `report_spec.md`로 쓴 뒤 스크립트만 실행한다. 스크립트를 고치지 않으므로 장표 단계에 토큰이 거의 들지 않는다.
  ```bash
  python "<RM_DIR>/scripts/spec_to_company.py" "<작업폴더>/report_spec.md" -o "<작업폴더>/<제목>.pptx"            # 가로(row)
  python "<RM_DIR>/scripts/spec_to_company.py" "<작업폴더>/report_spec.md" -o "<작업폴더>/<제목>.pptx" --layout col  # A4 세로
  ```
  - 가로: `## 장` 하나 = 장표 한 장. 표지는 row.pptx 1쪽(제목·목차·날짜·부서)을 채운다. HL은 제목 바 아래 가운데 헤드라인, 【제목】은 라벨 바 + 테두리 박스, `"문장"`은 하단 17pt 남색 문장.
  - 세로: 배너 아래로 `## 장`을 '1. 장 제목 : 헤드라인' 섹션으로 쌓고 쪽이 차면 넘긴다(`## [새쪽] …`은 새 쪽에서 시작). 【라벨】 바로 아래 줄은 회색 둥근 라벨 + 본문 행. 예시 문안: `examples/meeting_col_spec.md`.
  - 머리말에 `layout: col`을 적어도 된다. 실행 후 '넘칠 수 있음' 경고가 나오면 문안을 줄이거나 장을 나눈다. 생성 후 LibreOffice 등으로 렌더링해 넘침·겹침을 확인한다.
- **스크립트로 안 되는 배치만 직접 그린다:** 원형 라벨, 2단 박스 등 특별한 배치가 꼭 필요할 때만 `examples/onepage_vertical_meeting/build_onepage.py`(세로 = col.pptx) 또는 `examples/onepage_landscape_privacy/build_onepage.py`(가로 = row.pptx)를 작업 폴더로 복사해 `main()` 안의 문안을 바꿔 실행한다(토큰이 더 든다). 결과 파일 경로는 첫 인자로 넘긴다.
- **도형양식(templates/bullet.pptx):** 가로양식과 같은 크기(10.83x7.5in)·마스터의 도형·레이아웃 견본 모음(둥근 사각형 나열, 인용+근거 박스, 단계 흐름, 일정표, 표·다이어그램 등)이다. 가로양식 장표에 구조도·흐름도가 필요하면 먼저 이 파일에서 맞는 견본 장을 골라 도형 배치·색·크기를 따르거나 python-pptx로 해당 도형을 복사해 문안만 바꾼다. 이걸로 안 되는 복잡한 시각화만 PPT Master(T2)로 넘긴다.
- 필요 패키지: `python-pptx`.
- **koreport 하위 스킬 파일 이름:** 업로드 규칙(ZIP 안 SKILL.md는 1개만)에 맞추려고 `koreport/GUIDE.md`, `koreport/skills/<이름>/GUIDE.md`로 바꿨다. 원본 문서에서 `SKILL.md`라고 부르는 하위 파일은 모두 이 `GUIDE.md`를 뜻한다.

경로 표기: 이 파일이 있는 폴더를 `RM_DIR`로 잡고 모든 경로를 그 아래 절대경로로 쓴다. `cd`로 하위 폴더에 들어가지 않는다.

## 사람과 AI의 역할 (koreport 원칙 그대로)

- 독자의 니즈·보고 방향·핵심 메시지의 최종 선택은 사람이 한다. 주어지지 않았으면 지어내지 않고, 방향을 바꿀 정보만 짧게 묻거나 가정을 밝힌 뼈대를 먼저 보여 준다.
- 수치·사례·출처·일정·담당자를 창작하지 않는다. 없으면 `확인 필요`로 둔다(점검기가 잡아낸다).
- 만들지 않은 파일을 만들었다고 말하지 않는다.

## 전체 흐름

| 단계 | 하는 일 | 읽을 파일 | 사용자 확인 |
|---|---|---|---|
| **A1. 뼈대** | 목적 문장, 목차, 장별 헤드라인만 작성 | `workflows/01-report-spec.md` §1~3 | **필수** — 뼈대를 보여 주고 방향 확인 |
| **A2. 문안** | 장별 본문을 `report_spec.md` 문법으로 작성 | `01` §4~5, `02-spec-format.md` | 필요 시 |
| **A3. 점검** | `spec_to_pptx.py --check` + koreport 퇴고 | `01` §6 | — |
| **B. 장표** | 출력 단계 선택(T0~T3) 후 생성 | `workflows/03-pptx-output.md` | 단계를 사용자가 정하지 않았으면 T1(회사 양식)로 진행 |

**순서를 지킨다.** A1 뼈대가 확인되기 전에 본문을 쓰지 않고, A2 문안이 확정되기 전에 장표를 만들지 않는다. 장표를 먼저 그리면 수정 때마다 장표를 다시 만들어 비용이 몇 배로 든다.

## 출력 단계 (자세히는 03-pptx-output.md)

| 단계 | 결과물 | 비용 | 언제 |
|---|---|---|---|
| **T0** 문안만 | `report_spec.md` | 최소 | 문안 검토·결재용 초안, 사내 양식에 직접 옮길 때 |
| **T1** 기본 장표 *(기본값)* | 회사 양식(row/col) PPTX (`spec_to_company.py` 렌더링) | 낮음 — 렌더링에 LLM 토큰 없음 | 대부분의 사내 보고 |
| **T2** 기본 + 시각화 일부 | T1 + `@ppt-master` 장만 PPT Master Quick으로 그려 병합 | 중간 — 시각화 장 수에 비례 | 개념도·프로세스·비교 도식이 메시지를 결정하는 2~3장 |
| **T3** 전체 디자인 | PPT Master 전체 실행(문안은 확정본으로 전달) | 높음 | 대외 발표·행사용, 사용자가 명시적으로 요청할 때만 |

T2·T3에서 PPT Master를 실행할 때는 `ppt-master/SKILL.md`의 규칙(무결성 검사, 라우팅, 게이트)을 그대로 따르고, 그 `SKILL_DIR`은 `RM_DIR/ppt-master`다.

## 요청 유형별 처리

- **"보고서를 PPT로 만들어 줘" (자료·주제만 있음):** A1 → 확인 → A2 → A3 → T1(회사 양식, 요청 없이도 적용). 시각화가 필요해 보이는 장은 `@ppt-master` 후보로 제안만 하고, T2 진행은 사용자에게 맡긴다.
- **이미 쓴 문안·보고서 초안이 있음:** 요청이 "교정"뿐이어도 koreport 진단(`koreport/references/12-revision.md`)으로 먹통/불통부터 가려 문안을 교정한다. 결과물이 PPT(또는 PPT 초안을 준 경우)면 `report_spec.md`로 옮겨 A3 → B(회사 양식)로 교정본 장표까지 만든다.
- **"빨리·싸게":** 뼈대 확인을 한 번만 받고 T1로 끝낸다.
- **"예쁘게·디자인 중심":** 문안은 A단계로 확정하고 T2(핵심 장만) 또는 T3. T3는 비용이 크다는 점을 한 줄로 알린다.
- **기존 PPTX 수정:** 문구만 바꾸면 해당 장의 문안을 고쳐 회사 양식으로 다시 렌더링한다(T1 덱). PPT Master로 만든 장이나 외부 PPTX는 `ppt-master`의 Edit Native PPTX 경로를 쓴다.
- **문안 교정만, 메일, 주간업무:** PPT가 필요 없으므로 `koreport/GUIDE.md`와 해당 전용 스킬만 쓴다.

## 우선순위

1. 사용자가 지정한 목적·분량·형식, 회사 양식과 표준 용어가 이 스킬의 기준보다 우선한다.
2. 문안 규칙은 koreport가, 장표 시각 설계 규칙은 PPT Master가 소유한다. 이 스킬은 둘을 잇는 순서와 형식(`report_spec.md`)만 정한다.
3. PPT Master 장표에서도 확정 문안(헤드라인·수치·용어)은 바꾸지 않는다. 바꿀 필요가 생기면 `report_spec.md`를 고치고 다시 내려보낸다.
