# 자치법규 정비 후보 수집 공유 레포

이 레포는 자치법규정보시스템(ELIS)의 현행 자치법규를 수집하고, 통합·명칭변경·조직개편 같은 상황에서 검토가 필요한 조문과 별지·서식을 빠르게 찾아내는 학습용 예제입니다.

개발자용 제품이라기보다, 비개발자 행정직 공무원이 "이런 방식으로 AI와 스크립트를 나눠 쓰면 대량 자치법규 점검을 빠르게 할 수 있구나"를 이해할 수 있도록 정리했습니다.

## 이 프로젝트가 한 일

사람이 자치법규 사이트에서 하나씩 열어보면 며칠 이상 걸릴 수 있는 일을 다음 흐름으로 나누었습니다.

1. 대상 지자체와 ELIS 조회 코드를 정리한다.
2. 스크립트가 각 지자체의 현행 자치법규 목록, 본문, 별지·별표 첨부파일을 수집한다.
3. 스크립트가 `전라남도`, `도지사`, `광주광역시장`, `시 조례`, `도비` 같은 검토 후보 표현을 대량 검색한다.
4. Codex Agent가 결과를 보고 오탐 패턴, 서식 문제, 누락 가능성을 찾아 스크립트를 보정한다.
5. 사람이 최종 업무 방향과 판단 기준을 정하고, 제출 전에는 실제 정비대상인지 검수한다.

자동화는 "정답 확정"이 아니라 "검토해야 할 후보를 빠짐없이, 근거와 함께 모아주는 역할"을 맡습니다.

## 역할 분담

| 주체 | 맡은 일 | 이유 |
| --- | --- | --- |
| 사람 | 공문 요구사항 해석, 검색 기준 결정, 오탐 판단, 최종 제출 여부 결정 | 행정 맥락과 책임 있는 판단은 사람이 해야 함 |
| Codex Agent | 작업 순서 설계, 스크립트 수정, 오류 원인 분석, 산출물 검증, 설명자료 작성 보조 | 코드와 문서 사이를 오가며 반복 보정하기 좋음 |
| 자동화 스크립트 | 목록 수집, 본문·별지 다운로드, 텍스트 추출, 후보 검색, 엑셀·문서 생성 | 사람이 반복하기 어려운 대량 처리를 빠르고 일관되게 수행 |

핵심은 셋 중 하나가 전부 하는 것이 아닙니다. 사람은 "무엇을 봐야 하는지"를 정하고, 스크립트는 "전부 찾아서 정리"하고, Codex Agent는 중간에서 스크립트를 업무 목적에 맞게 계속 고쳐 줍니다.

## 폴더 구성

| 경로 | 내용 |
| --- | --- |
| `data/targets_sample.csv` | 전남 22개 시군 + 광주 5개 자치구 예시 대상목록 |
| `scripts/verify_elis_targets.py` | 대상 코드가 ELIS에서 실제 응답하는지 확인 |
| `scripts/collect_elis_fulltext.py` | 자치법규 본문과 별지·별표 원천자료 수집 |
| `scripts/run_body_delivery_all.py` | 본문 조문 기준 후보 검색 및 제출서식 생성 |
| `scripts/scan_attachment_delivery_all.py` | 별지·별표 추출텍스트 기준 후보 검색 |
| `scripts/merge_body_attachment_delivery.py` | 본문 후보와 별지 후보를 통합 |
| `docs/project_explanation_for_admins.md` | 비개발자 행정직용 프로젝트 설명서 |
| `docs/technical_appendix_for_ai.md` | 상대편 AI가 이해할 수 있는 기술 붙임자료 |
| `docs/workflow_roles.md` | 사람·Codex Agent·스크립트 역할 분담 설명 |
| `docs/search_criteria.md` | 후보 검색 기준 예시 |

## 빠른 실행 예시

Python 3.11 이상을 권장합니다.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 법제처 OC가 필요한가?

이 레포의 기본 스크립트는 법제처 Open API가 아니라 자치법규정보시스템(ELIS) 웹 화면을 조회합니다. 그래서 `OC` 인증키 없이 실행할 수 있습니다.

다만 별도로 법제처 Open API, Korean Law MCP, 다른 법령 API 도구를 함께 붙여서 쓰려면 그 도구 쪽에서 발급받은 `OC`가 필요할 수 있습니다. 이 레포의 스크립트에는 OC 입력칸이나 환경변수가 없습니다.

## 별지·서식 파일은 어떻게 읽었나?

ELIS에서 내려받은 별지·별표 첨부파일은 확장자와 파일 헤더를 보고 텍스트를 추출합니다.

- 구형 `.hwp`: `olefile`로 HWP5 OLE 저장소를 열고 `BodyText/Section*` 스트림을 직접 읽어 텍스트를 뽑습니다.
- `.hwpx`, `.docx`, `.xlsx`: 내부가 ZIP 구조라 `zipfile`로 XML을 열어 텍스트를 뽑습니다.
- `.pdf`: `pypdf`로 페이지 텍스트를 읽습니다.

전용 상용 변환기를 쓰는 방식은 아니어서, 일부 HWP 서식이나 복잡한 표는 깨지거나 누락될 수 있습니다. 그래서 결과는 최종 정답이 아니라 사람이 검토할 후보 자료로 봐야 합니다.

대상 코드 확인:

```powershell
python .\scripts\verify_elis_targets.py --targets .\data\targets_sample.csv
```

일부 지자체만 원천수집:

```powershell
python .\scripts\collect_elis_fulltext.py --entities-csv .\data\targets_sample.csv --only "여수시,광주광역시 남구"
```

본문 기준 후보 생성:

```powershell
python .\scripts\run_body_delivery_all.py --targets .\data\targets_sample.csv --only "여수시,광주광역시 남구"
```

별지·서식 후보 생성:

```powershell
python .\scripts\scan_attachment_delivery_all.py --targets .\data\targets_sample.csv --source-date 20260511 --only "여수시,광주광역시 남구"
```

본문+별지 통합:

```powershell
python .\scripts\merge_body_attachment_delivery.py --targets .\data\targets_sample.csv --body-date 20260511 --attachment-date 20260511 --only "여수시,광주광역시 남구"
```

## 공유 시 주의

이 레포에는 대용량 원천수집 결과와 실제 최종 제출파일을 포함하지 않습니다. 실행하면 `02_원천수집`, `04_제출서식`, `05_총괄보고` 같은 작업 폴더가 생기며, 이들은 `.gitignore`로 제외됩니다.

ELIS 화면 구조가 바뀌면 수집 스크립트의 HTML 파싱 부분을 보정해야 할 수 있습니다. 또한 자동 검색 결과는 후보일 뿐이며, 실제 정비대상 여부는 담당자가 조문 맥락과 소관부서 의견을 확인해야 합니다.
