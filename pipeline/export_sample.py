"""
지금까지 지오코딩된 데이터 중 일부(N건)만 뽑아서 검토용 엑셀로 내보낸다.
build_dataset.py 와 동일한 조인 로직을 재사용한다.

실행: python pipeline/export_sample.py [N건수, 기본 100]
"""
import sys

import pandas as pd

from build_dataset import (
    DB_SHEET,
    DB_SOURCE,
    NEEDED_COLUMNS,
    count_units_by_address,
    load_addr_map,
    load_cache,
    load_project_no_map,
    normalize_key,
    pad_elevator_no,
    pad_project_no,
)
import pyxlsb
import os

OUT_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "sample_100.xlsx")


def main():
    limit = int(sys.argv[1]) if len(sys.argv) > 1 else 100

    cache = load_cache()
    addr_map = load_addr_map()
    project_map = load_project_no_map()
    unit_counts = count_units_by_address(addr_map)
    print(f"지오코딩 캐시: {len(cache)}건 / 주소 매핑: {len(addr_map)}건 / 프로젝트번호 매핑: {len(project_map)}건")

    records = []
    with pyxlsb.open_workbook(DB_SOURCE) as wb:
        with wb.get_sheet(DB_SHEET) as sheet:
            rows = sheet.rows()
            header = [c.v for c in next(rows)]
            idx = {name: header.index(name) for name in NEEDED_COLUMNS}
            for row in rows:
                if len(records) >= limit:
                    break
                vals = [c.v for c in row]
                eno = normalize_key(vals[idx["ELEVATORNO"]])
                addr = addr_map.get(eno)
                if not addr:
                    continue
                coord = cache.get(addr)
                if not coord:
                    continue
                name = vals[idx["BULDNM"]]
                records.append(
                    {
                        "승강기번호": pad_elevator_no(eno),
                        "위도": coord["lat"],
                        "경도": coord["lng"],
                        "건물명": name,
                        "주소": addr,
                        "제조업체": vals[idx["MANUFACTURERNAME"]],
                        "유지보수업체": vals[idx["MNTCPNYNM"]],
                        "대수": unit_counts.get(addr),
                        "기종": vals[idx["ELVTRMODEL"]],
                        "상태": vals[idx["ELVTRSTTS"]],
                        "종류": vals[idx["ELVTRKINDNM"]],
                        "최초설치일": vals[idx["FRSTINSTALLATIONDE"]],
                        "설치일": vals[idx["INSTALLATIONDE"]],
                        "프로젝트번호": pad_project_no(project_map.get(eno)),
                    }
                )

    df = pd.DataFrame(records)
    df.to_excel(OUT_FILE, index=False)
    print(f"완료: {len(records)}건 -> {OUT_FILE}")


if __name__ == "__main__":
    main()
