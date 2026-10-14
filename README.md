# koreport-skill

[![skills.sh](https://skills.sh/b/iksangyoo-sys/koreport-skill)](https://skills.sh/iksangyoo-sys/koreport-skill)

유익상의 「AI가 못 쓰는 보고서 쓰기」 교안을 담은 에이전트 스킬입니다. 한국어 업무 보고서를 기획·작성·진단·수정할 때 쓰며, 한국 기업 보고서 문장 패턴 22개(13편), 보고서 어휘 240묶음(14편), 실제 사내 보고서 장표에서 정리한 1-page 장표 양식(15편)과 교정 사례집(16편), 학습자용 「보고서 문장 공식」 도구가 함께 들어 있습니다.

## 설치

### skills.sh (Claude Code 등 코딩 에이전트)

```bash
npx skills add iksangyoo-sys/koreport-skill
```

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
| `references/` | 교안 16편 (13편 = 문장 패턴 공식, 14편 = 보고서 어휘, 15편 = 1-page 장표 양식, 16편 = 교정 사례집) |
| `assets/sentence-patterns.json` | 문장 패턴 원천 데이터 |
| `assets/report-vocabulary.json` | 보고서 어휘 원천 데이터 (240묶음, 2,452단어) |
| `assets/sentence-patterns-guide.md` | 학습자 사용안내서 원문 |
| `assets/sentence-patterns.html` | 학습자용 「보고서 문장 공식」 (빌드 결과) |
| `scripts/build_patterns.py` | 위 데이터로 13편·14편과 HTML을 다시 만드는 스크립트 |

패턴·어휘·안내서를 고친 뒤에는 `python scripts/build_patterns.py`를 실행해 `references/13-sentence-patterns.md`, `references/14-report-vocabulary.md`, `assets/sentence-patterns.html`을 다시 만듭니다.
