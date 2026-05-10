# 실행 순서

## 1. 대상목록 준비

`data/targets_sample.csv`를 복사해 필요한 지자체만 남기거나, 새 CSV를 만듭니다.

필수 컬럼은 다음 네 개입니다.

```csv
권역,지자체명,ctpvCd,sggCd
```

기준 수량 비교를 하고 싶으면 `예상건수` 또는 `ELIS_목록건수_YYYYMMDD` 컬럼을 추가합니다.

## 2. 코드 검증

```powershell
python .\scripts\verify_elis_targets.py --targets .\data\targets_sample.csv
```

## 3. 원천자료 수집

처음에는 한두 곳만 시험합니다.

```powershell
$env:ELIS_COLLECT_DATE="20260511"
python .\scripts\collect_elis_fulltext.py --entities-csv .\data\targets_sample.csv --only "여수시"
```

## 4. 본문 후보 검색

```powershell
python .\scripts\run_body_delivery_all.py --targets .\data\targets_sample.csv --only "여수시"
```

## 5. 별지·서식 후보 검색

```powershell
python .\scripts\scan_attachment_delivery_all.py --targets .\data\targets_sample.csv --source-date 20260511 --only "여수시"
```

## 6. 통합본 생성

```powershell
python .\scripts\merge_body_attachment_delivery.py --targets .\data\targets_sample.csv --body-date 20260511 --attachment-date 20260511 --only "여수시"
```

## 7. 검수

엑셀에서 `제출서식` 시트를 먼저 봅니다. 다음 항목을 확인합니다.

- 조문 위치가 비어 있거나 이상하지 않은지
- 원문 내용이 중간에서 잘리지 않았는지
- 단순 주소·기관명 반복이 너무 많이 잡히지 않았는지
- 실제 정비 필요성이 있는지
- 소관부서에 넘길 때 설명이 충분한지
