# koreport-skill

[![skills.sh](https://skills.sh/b/iksangyoo-sys/koreport-skill)](https://skills.sh/iksangyoo-sys/koreport-skill)

유익상의 「AI가 못 쓰는 보고서 쓰기」 교안을 담은 에이전트 스킬입니다. 한국어 업무 보고서를 기획·작성·진단·수정할 때 쓰며, 학습자용 「보고서 문장 공식」 도구가 함께 들어 있습니다.

`skills/` 에는 전용 스킬 9개가 있습니다. 보고서 유형별 6개(문제해결·추진방안 보고서, QSS 보고서, 업무 이메일, 주간업무 보고, 회의결과 보고, 1-page 장표 양식)와 작업별 3개(문장 패턴 공식 22개, 보고서 어휘 240묶음, 교정 사례집)입니다. 각 스킬은 `koreport-skill`의 일반 방법론 위에 해당 유형·작업에 특화된 내용만 얹습니다. 전용 스킬을 독립 스킬로 설치하는 방법은 아래 「설치」에 있습니다.

## 설치

### skills.sh (Claude Code 등 코딩 에이전트)

```bash
npx skills add iksangyoo-sys/koreport-skill --full-depth
```

- `--full-depth`를 붙여야 메인 스킬과 `skills/` 의 전용 스킬 9개가 모두 독립 스킬로 잡힙니다. 붙이지 않으면 메인 스킬(`koreport-skill`) 하나만 설치되고, 전용 스킬은 메인 스킬이 필요할 때 읽는 파일로만 쓰입니다.
- 설치할 스킬 목록 보기: `npx skills add iksangyoo-sys/koreport-skill --full-depth --list`
- 전용 스킬 하나만 설치: `npx skills add iksangyoo-sys/koreport-skill --full-depth --skill weekly-report`
- 이미 설치했다면 `npx skills update`로 최신 버전을 받습니다.
- Claude Code에만 설치: `npx skills add iksangyoo-sys/koreport-skill -a claude-code`
- 모든 프로젝트에서 쓰기(사용자 전체 설치): 뒤에 `-g`를 붙입니다.
- Node.js가 설치되어 있어야 합니다.

### claude.ai (웹·데스크톱 앱)

1. 이 저장소의 폴더를 `koreport-skill` 폴더째 zip으로 묶습니다.
2. claude.ai의 **Customize → Skills**에서 **+ → Create skill → 업로드**로 zip을 올립니다.
3. 예전 버전이 있으면 끄거나 지웁니다.

사용자 지정 스킬은 Pro·Max·Team·Enterprise 요금제에서 코드 실행 기능을 켜야 쓸 수 있습니다.

## 학습자용 도구

웹에서 바로 쓰기: https://iksangyoo-sys.github.io/koreport-skill/

`assets/sentence-patterns.html` 파일 하나를 더블클릭하면 「보고서 문장 공식」이 열립니다. 인터넷 없이 동작하며, 화면 오른쪽 위 **사용안내서** 버튼에 사용법이 있습니다.

## 구성

| 경로 | 내용 |
| --- | --- |
| `SKILL.md` | 스킬 본문 — 작성 절차와 작업별로 읽을 파일 |
| `references/` | 교안 12편 (01~12편) |
| `skills/problem-action-report/` | 전용 스킬 — 문제해결형·추진방안형 보고서 |
| `skills/qss-report/` | 전용 스킬 — QSS(현장 개선활동) 보고서 |
| `skills/business-email/` | 전용 스킬 — 업무 이메일 |
| `skills/weekly-report/` | 전용 스킬 — 주간업무 실적·계획 보고 |
| `skills/meeting-report/` | 전용 스킬 — 회의결과 보고 |
| `skills/onepage-format/` | 전용 스킬 — 한국 기업 1-page 장표 양식 |
| `skills/sentence-patterns/` | 전용 스킬 — 문장 패턴 공식 22개 (빌드 결과) |
| `skills/report-vocabulary/` | 전용 스킬 — 보고서 어휘 240묶음, 2,452단어 (빌드 결과) |
| `skills/correction-casebook/` | 전용 스킬 — 교정·교열·윤문 사례집 |
| `assets/sentence-patterns.json` | 문장 패턴 원천 데이터 |
| `assets/report-vocabulary.json` | 보고서 어휘 원천 데이터 (240묶음, 2,452단어) |
| `assets/sentence-patterns-guide.md` | 학습자 사용안내서 원문 |
| `assets/sentence-patterns.html` | 학습자용 「보고서 문장 공식」 (빌드 결과) |
| `scripts/build_patterns.py` | 위 데이터로 sentence-patterns·report-vocabulary 스킬과 HTML을 다시 만드는 스크립트 |

패턴·어휘·안내서를 고친 뒤에는 `python scripts/build_patterns.py`를 실행해 `skills/sentence-patterns/SKILL.md`, `skills/report-vocabulary/SKILL.md`, `assets/sentence-patterns.html`을 다시 만듭니다.
