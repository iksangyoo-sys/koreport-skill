# templates

`report_spec_template.md`(빈 문안 양식)만 들어 있습니다.

회사 장표 양식 파일은 회사 내부 자료라서 저장소에 올리지 않았습니다. 스킬을 쓰려면 각자의 양식을 아래 이름으로 이 폴더에 넣으세요.

| 파일 | 용도 |
|---|---|
| `row.pptx` | 가로 1-page 양식 (기본) |
| `col.pptx` | A4 세로 1장 양식 |
| `bullet.pptx` | 가로 도형·레이아웃 견본 |

다른 이름의 양식을 쓰려면 `scripts/spec_to_company.py --template <파일>`로 직접 지정합니다. 양식 없이도 `scripts/spec_to_pptx.py`(고정 블루 양식)는 동작합니다.
