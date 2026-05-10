from __future__ import annotations

import argparse
import csv
import datetime as dt
import re
import time
from pathlib import Path

import requests
from bs4 import BeautifulSoup


BASE_URL = "https://www.elis.go.kr"


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
            )
        }
    )
    return session


def get_csrf(session: requests.Session, ctpv_cd: str, sgg_cd: str) -> tuple[str, str]:
    referer = f"{BASE_URL}/allalr/allAlrList?ctpvCd={ctpv_cd}&sggCd={sgg_cd}"
    response = session.get(referer, timeout=30)
    response.raise_for_status()
    match = re.search(r'name="_csrf" content="([^"]+)"', response.text)
    if not match:
        match = re.search(r'name="_csrf" value="([^"]+)"', response.text)
    if not match:
        raise RuntimeError("CSRF token not found")
    return match.group(1), referer


def verify_one(session: requests.Session, row: dict[str, str]) -> dict[str, str]:
    ctpv_cd = row["ctpvCd"].strip()
    sgg_cd = row["sggCd"].strip()
    result = dict(row)
    result["검증상태"] = "OK"
    result["응답건수"] = ""
    result["첫번째법규"] = ""
    result["오류"] = ""

    try:
        csrf, referer = get_csrf(session, ctpv_cd, sgg_cd)
        data = {
            "_csrf": csrf,
            "curPage": "1",
            "sortSe": "",
            "sortType": "",
            "srchTabSeCd": "",
            "ctpvCd": ctpv_cd,
            "sggCd": sgg_cd,
            "ordnRuleSe": "",
            "srchKwd": "",
            "pageSize": "10",
        }
        response = session.post(
            f"{BASE_URL}/allalr/allAlrList",
            data=data,
            headers={"X-CSRF-TOKEN": csrf, "Referer": referer},
            timeout=30,
        )
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
        total_node = soup.select_one(".list-top .info span")
        first_node = soup.select_one("div.col-table.only-pc tbody tr td:nth-of-type(2)")
        result["응답건수"] = total_node.get_text(strip=True).replace(",", "") if total_node else ""
        result["첫번째법규"] = first_node.get_text(" ", strip=True) if first_node else ""

        expected = (row.get("예상건수") or "").strip()
        if not expected:
            expected_columns = [key for key in row if key.startswith("ELIS_목록건수")]
            expected = (row.get(expected_columns[0]) or "").strip() if expected_columns else ""
        if expected and result["응답건수"] and expected != result["응답건수"]:
            result["검증상태"] = "COUNT_CHANGED"
        if result["응답건수"] == "0":
            result["검증상태"] = "ZERO"
    except Exception as exc:  # noqa: BLE001 - verification report should keep going.
        result["검증상태"] = "ERROR"
        result["오류"] = f"{type(exc).__name__}: {exc}"
    return result


def write_markdown(output_csv: Path, rows: list[dict[str, str]], md_path: Path) -> None:
    ok = sum(1 for row in rows if row["검증상태"] == "OK")
    changed = sum(1 for row in rows if row["검증상태"] == "COUNT_CHANGED")
    zero = sum(1 for row in rows if row["검증상태"] == "ZERO")
    error = sum(1 for row in rows if row["검증상태"] == "ERROR")
    total_count = sum(int(row["응답건수"]) for row in rows if row.get("응답건수", "").isdigit())

    lines = [
        "# ELIS 대상 지자체 검증 결과",
        "",
        f"- 검증일: {dt.date.today():%Y-%m-%d}",
        f"- 대상: {len(rows)}개 지자체",
        f"- OK: {ok}개",
        f"- 수량변동: {changed}개",
        f"- 0건: {zero}개",
        f"- 오류: {error}개",
        f"- 응답 총건수: {total_count:,}건",
        f"- 상세 CSV: `{output_csv.name}`",
        "",
        "## 지자체별 응답건수",
        "",
        "| 권역 | 지자체명 | ctpvCd | sggCd | 응답건수 | 상태 |",
        "| --- | --- | --- | --- | ---: | --- |",
    ]
    for row in rows:
        lines.append(
            f"| {row.get('권역','')} | {row.get('지자체명','')} | {row.get('ctpvCd','')} | "
            f"{row.get('sggCd','')} | {row.get('응답건수','')} | {row.get('검증상태','')} |"
        )
    md_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Verify ELIS target municipality codes.")
    parser.add_argument("--project-dir", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--targets", type=Path, default=None)
    parser.add_argument("--sleep", type=float, default=0.1)
    args = parser.parse_args()

    project_dir = args.project_dir
    targets = args.targets or project_dir / "data" / "targets_sample.csv"
    today = dt.date.today().strftime("%Y%m%d")
    output_csv = project_dir / "data" / f"target_entities_verified_{today}.csv"
    output_md = project_dir / "data" / f"target_entities_verified_{today}.md"
    output_csv.parent.mkdir(parents=True, exist_ok=True)

    with targets.open("r", encoding="utf-8-sig", newline="") as file:
        rows = list(csv.DictReader(file))

    session = make_session()
    verified: list[dict[str, str]] = []
    for row in rows:
        verified.append(verify_one(session, row))
        print(f"{verified[-1]['지자체명']}: {verified[-1]['응답건수']} {verified[-1]['검증상태']}")
        time.sleep(args.sleep)

    fieldnames = list(verified[0].keys()) if verified else []
    with output_csv.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(verified)
    write_markdown(output_csv, verified, output_md)

    print(output_csv)
    print(output_md)


if __name__ == "__main__":
    main()
