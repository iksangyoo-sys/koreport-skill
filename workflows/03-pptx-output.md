# B단계 — 장표 만들기 (필요한 만큼만)

전제: `report_spec.md`가 A단계 점검(`--check`)을 통과했다. 문안 수정은 항상 `report_spec.md`에서 하고 다시 렌더링한다.

## 단계 고르기

사용자가 정하지 않았으면 **T1(회사 양식)**로 진행한다. 회사 양식(templates의 row·col)은 사용자가 말하지 않아도 적용한다. **장표를 만들기 전에 가로(row)/세로(col)와 장 수를 묻는다**(요청에 이미 있으면 생략). T2·T3로 올릴 때는 늘어나는 비용을 한 줄로 알린다.

| 단계 | 언제 | LLM 비용 |
|---|---|---|
| T0 문안만 | 결재 전 문안 검토, 사내 양식 파일에 직접 옮길 때 | 렌더링 없음 |
| T1 기본 장표 | 대부분 | 스크립트 실행뿐 |
| T2 + 시각화 일부 | 도식이 메시지를 결정하는 장이 있을 때(권장 2~3장 이내) | 시각화 장 수 × PPT Master 장당 비용 |
| T3 전체 디자인 | 대외 발표·행사, 사용자가 명시적으로 요청 | 전체 장 × PPT Master 장당 비용 + 전략 단계 |

`@ppt-master` 후보 판단 기준: 프로세스·순환·인과 구조, 로드맵·간트, 조직·역할도, 수치 추세 차트, 개념 프레임(3대 방향 n개 과제 등). **글과 표로 충분한 장은 올리지 않는다.**

## T0 — 문안만

`report_spec.md`를 그대로 전달한다. 필요하면 Markdown 표·개조식 그대로 사내 양식에 붙여 넣을 수 있다고 안내한다.

## T1 — 기본 장표 (회사 양식)

기본은 templates의 회사 양식 견본 배치다. 문안을 넘겨 스크립트만 실행한다.

```bash
# 가로 = row.pptx 견본 배치(묶음 문장·회색선 박스·라벨 바·표·원형 라벨·하단 따옴표), 장마다 꽉 차게
python "<RM_DIR>/scripts/spec_to_onepage.py" "<project>/onepage_spec.md" -o "<project>/<제목>.pptx" --pages <장 수>
# 세로 = col.pptx 견본 배치(배너·제목 아래 '1. 섹션 : 헤드라인' + 둥근 라벨 행)
python "<RM_DIR>/scripts/spec_to_company.py" "<project>/report_spec.md" -o "<project>/<제목>.pptx" --layout col
```

- 가로 문법은 `02-spec-format.md`의 '가로 견본 배치 문법'과 `templates/onepage_spec_template.md`. 여러 장이면 `[새쪽] 제목`으로 나눈다. 가로 렌더러는 report_spec.md(`## 장`, `HL:`)도 읽어 HL을 ❍ 묶음 문장으로 쓴다.
- 세로 문법은 `02-spec-format.md` 그대로다. 【라벨】 아래 줄이 '둥근 라벨 + 본문' 행이 되고, `## [새쪽] …`으로 쪽을 나눈다.
- 표지가 붙은 긴 가로 보고서만 `spec_to_company.py`(row, `## 장`마다 1장)를 쓴다.
- 구조도·흐름도: `templates/bullet.pptx` 견본 도형을 복사해 쓴다.

### 고정 블루 양식 (사용자가 이름을 대어 요청할 때만)

```bash
python "<RM_DIR>/scripts/spec_to_pptx.py" "<project>/report_spec.md" -o "<project>/<제목>.pptx"
```

- 16:9, 맑은 고딕, 블루 단색 계열, 장 4층 구조(장 제목 바 → 헤드라인 → 묶음 본문 → ☞ → 하단 문장), 쪽 번호 '현재/전체'.
- 모든 글자는 편집 가능한 텍스트 상자, 표는 PowerPoint 표다.
- 출력 끝에 양식 점검 지적이 나오면 spec을 고쳐 다시 실행한다.
- 의존성: `python-pptx` (`pip install python-pptx`).

## T2 — 기본 장표 + 일부 장만 PPT Master

1. `report_spec.md`에서 대상 장 제목 끝에 `@ppt-master`를 붙인다.
2. T1을 실행해 기본 덱을 만든다(대상 장도 문안대로 들어가 있어 그대로 써도 된다).
3. 대상 장만 뽑아 PPT Master 브리프를 만든다.
   ```bash
   python "<RM_DIR>/scripts/spec_to_pptx.py" "<project>/report_spec.md" --extract-visual "<project>/visual_brief.md"
   ```
   브리프에는 확정 문안, 1-page 양식 색·글꼴 규칙, 기본 덱 쪽 ↔ 시각화 덱 쪽 대응표, 병합 명령이 들어 있다.
4. PPT Master를 **Quick 프로필**로 실행한다. `SKILL_DIR = <RM_DIR>/ppt-master`. 사용자의 요청은 "visual_brief.md로 빠르게(quick) 생성, 전략·확인 단계 생략"이다.
   - `ppt-master/SKILL.md`의 Mandatory Load Order(무결성 검사 → routing → quick-generate.md)를 따른다.
   - 소스는 `visual_brief.md` 하나다. 문안은 바꾸지 않는다.
   - 출력은 flat 구조, 16:9.
5. 결과를 기본 덱에 병합한다.
   ```bash
   python "<RM_DIR>/scripts/merge_slides.py" "<기본.pptx>" "<ppt-master 결과.pptx>" --map 4=1,6=2 -o "<최종.pptx>"
   ```
   `--map`은 브리프의 대응표 값을 그대로 쓴다. 이미지·차트·노트·전환을 함께 옮긴다.

## T3 — 전체 PPT Master

- PPT Master의 Default 또는 Quick 경로로 실행하되, 소스로 `report_spec.md`를 주고 "문안(헤드라인·본문·수치)은 확정본이므로 바꾸지 말 것, 장 수·순서 동일"을 요구사항으로 명시한다. 전략 단계가 문안을 다시 짜지 않게 한다.
- Default 경로의 Stage-1 확인 게이트는 PPT Master 규칙대로 사용자 확인을 받는다.
- 1-page 양식을 유지하려면 `visual_brief.md`의 색·글꼴 규칙을 요구사항에 함께 넣는다.

## 수정이 생기면

| 수정 | 하는 일 |
|---|---|
| 문구·수치 | `report_spec.md` 수정 → T1 재실행 → (T2였다면) 병합만 재실행. 시각화 장의 문구가 바뀌었으면 그 장만 PPT Master로 다시 |
| 장 추가·순서 변경 | spec 수정 → T1 재실행 → `--extract-visual`로 대응표를 다시 뽑아 병합 |
| 시각화 장 디자인 | PPT Master 프로젝트에서 해당 장만 수정 후 병합 재실행 |

## 비용을 더 줄이는 습관

- 뼈대(A1) 확인 전에는 본문을 쓰지 않는다. 방향이 틀리면 전부 버려진다.
- PPT Master 문서는 실행할 때만, 선택된 경로 파일만 읽는다(SKILL.md 규칙).
- koreport의 큰 파일(`report-vocabulary`, `sentence-patterns`)은 검색으로 필요한 항목만 본다.
- 결과 확인은 PPTX를 이미지로 바꿔 몇 장만 본다. 전부 보지 않는다.
