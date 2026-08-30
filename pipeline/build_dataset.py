"""
국가DB(aa 시트) + 주소 매핑 + 지오코딩 캐시를 합쳐서
지도에서 바로 쓸 수 있는 data/elevators.json 을 만든다.

- 아직 지오코딩이 끝나지 않은 주소는 건너뛴다 (geocode.py를 며칠에 걸쳐 실행하면서
  이 스크립트를 다시 돌리면 점점 더 많은 데이터가 채워진다).
- 실행: python pipeline/build_dataset.py
"""
import json
import os

import pandas as pd
import pyxlsb

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE_FILE = os.path.join(BASE_DIR, "pipeline", "cache", "geocode_cache.json")
ADDRESS_SOURCE = os.path.join(BASE_DIR, "국가DB_SID_주소.xlsx")
DB_SOURCE = os.path.join(BASE_DIR, "국가DB_260805.xlsb")
DB_SHEET = "aa"
OUT_FILE = os.path.join(BASE_DIR, "data", "elevators.json")

NEEDED_COLUMNS = ["ELEVATORNO", "BULDNM", "MNTCPNYNM", "ELVTRSTTS", "ELVTRKINDNM", "INSTALLATIONDE"]


def load_cache():
    if not os.path.exists(CACHE_FILE):
        raise SystemExit(f"{CACHE_FILE} 이 없습니다. 먼저 pipeline/geocode.py를 실행하세요.")
    with open(CACHE_FILE, encoding="utf-8") as f:
        return json.load(f)


def load_addr_map():
    df = pd.read_excel(ADDRESS_SOURCE)
    df["ADDRESS1"] = df["ADDRESS1"].astype(str).str.strip()
    return dict(zip(df["ELEVATORNO"].astype(str), df["ADDRESS1"]))


def main():
    cache = load_cache()
    addr_map = load_addr_map()
    print(f"지오코딩 캐시: {len(cache)}건 / 주소 매핑: {len(addr_map)}건")

    records = []
    skipped_no_addr = 0
    skipped_no_coord = 0

    with pyxlsb.open_workbook(DB_SOURCE) as wb:
        with wb.get_sheet(DB_SHEET) as sheet:
            rows = sheet.rows()
            header = [c.v for c in next(rows)]
            idx = {name: header.index(name) for name in NEEDED_COLUMNS}
            for row in rows:
                vals = [c.v for c in row]
                eno = str(vals[idx["ELEVATORNO"]])
                addr = addr_map.get(eno)
                if not addr:
                    skipped_no_addr += 1
                    continue
                coord = cache.get(addr)
                if not coord:
                    skipped_no_coord += 1
                    continue
                records.append(
                    {
                        "id": eno,
                        "lat": coord["lat"],
                        "lng": coord["lng"],
                        "name": vals[idx["BULDNM"]],
                        "address": addr,
                        "company": vals[idx["MNTCPNYNM"]],
                        "status": vals[idx["ELVTRSTTS"]],
                        "kind": vals[idx["ELVTRKINDNM"]],
                        "installDate": vals[idx["INSTALLATIONDE"]],
                    }
                )

    os.makedirs(os.path.dirname(OUT_FILE), exist_ok=True)
    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False)

    print(f"완료: {len(records)}건 저장 -> {OUT_FILE}")
    print(f"주소 매핑 없음(스킵): {skipped_no_addr}건")
    print(f"좌표 미완료(아직 지오코딩 안됨, 스킵): {skipped_no_coord}건")


if __name__ == "__main__":
    main()
