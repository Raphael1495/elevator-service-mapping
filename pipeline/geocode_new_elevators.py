"""
신규 승강기만 좌표 변환 (국가DB_261105 기준).

- 메인 데이터: 국가DB_261105.xlsb (ELEVATORNO, ADDRESS1 포함)
- 비교 기준: 국가DB_SID_주소_260805.xlsx (직전 DB의 승강기번호 목록)
- 직전 DB에 없던 승강기만 대상으로 하고, 기존 승강기는 행정구역 개편으로 주소 표기가
  바뀌었어도 기존 주소/좌표를 그대로 쓴다(재변환 안 함).
- 이미 geocode_cache.json에 좌표가 있는 주소는 건너뛰고, 이전 실패 기록(geocode_failed.json)
  주소는 카카오 DB 갱신 가능성이 있어 한 번 더 시도한다(성공하면 실패 기록에서 제거).
- 결과: geocode_cache.json 갱신, 신규 승강기 목록 pipeline/cache/new_elevators_261105.json 저장.
- 실행: python pipeline/geocode_new_elevators.py
"""
import json
import os
import sys
import time

import pandas as pd
import pyxlsb

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from pipeline.build_dataset import normalize_key  # noqa: E402
from pipeline.geocode import (  # noqa: E402
    CACHE_FILE, FAILED_FILE, SLEEP_SEC, geocode_one, load_config, load_json, save_json,
)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
NEW_DB = os.path.join(BASE_DIR, "국가DB_261105.xlsb")
NEW_DB_SHEET = 1  # 시트명이 바뀌어도(승강기정보센터DB) 첫 시트를 쓴다
OLD_ADDR = os.path.join(BASE_DIR, "국가DB_SID_주소_260805.xlsx")
OUT_FILE = os.path.join(BASE_DIR, "pipeline", "cache", "new_elevators_261105.json")


def read_new_db():
    rows = {}
    with pyxlsb.open_workbook(NEW_DB) as wb:
        with wb.get_sheet(NEW_DB_SHEET) as sh:
            it = sh.rows()
            h = [c.v for c in next(it)]
            i, a = h.index("ELEVATORNO"), h.index("ADDRESS1")
            for r in it:
                v = [c.v for c in r]
                v += [None] * (len(h) - len(v))
                k = normalize_key(v[i])
                if k and v[a] is not None:
                    rows[k] = str(v[a]).strip()
    return rows


def main():
    rest_key = load_config()["kakao_rest_api_key"]
    new = read_new_db()
    old_keys = set(pd.read_excel(OLD_ADDR)["ELEVATORNO"].map(normalize_key))
    new_only = {k: a for k, a in new.items() if k not in old_keys}
    print(f"새 DB {len(new)}대 / 기존 DB 승강기 {len(old_keys)}대 / 신규 승강기 {len(new_only)}대")

    cache = load_json(CACHE_FILE, {})
    failed = load_json(FAILED_FILE, {})
    todo = sorted({a for a in new_only.values() if a not in cache})
    retry = sum(1 for a in todo if a in failed)
    print(f"변환 필요 주소 {len(todo)}개 (그중 이전 실패 기록 {retry}개 재시도)")

    ok = fail = 0
    still_failed = []
    for n, addr in enumerate(todo, 1):
        try:
            res = geocode_one(addr, rest_key)
        except Exception as e:  # 네트워크 오류는 실패로 기록하지 않고 건너뜀
            print(f"[오류] {addr}: {e}")
            continue
        if res:
            cache[addr] = res
            failed.pop(addr, None)
            ok += 1
        else:
            failed[addr] = True
            still_failed.append(addr)
            fail += 1
        if n % 200 == 0:
            save_json(CACHE_FILE, cache)
            save_json(FAILED_FILE, failed)
            print(f"  진행 {n}/{len(todo)} (성공 {ok}, 실패 {fail})")
        time.sleep(SLEEP_SEC)

    save_json(CACHE_FILE, cache)
    save_json(FAILED_FILE, failed)
    save_json(OUT_FILE, {"newElevators": new_only, "stillFailed": still_failed})
    print(f"완료: 성공 {ok}, 실패 {fail} / 캐시 {len(cache)}건, 실패기록 {len(failed)}건")


if __name__ == "__main__":
    main()
