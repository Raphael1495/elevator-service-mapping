"""
주소 -> 건축물대장 조회키(시군구코드 / 법정동코드 / 대지구분 / 본번 / 부번) 변환.

왜 필요한가:
  국가DB의 주소(ADDRESS1)는 전부 도로명주소인데, 건축물대장 API는 도로명주소로는
  조회할 수 없고 "지번" 기반 키(sigunguCd + bjdongCd + platGbCd + bun + ji)만 받는다.
  그래서 도로명주소 -> 지번 변환을 한 번 해두고 그 결과를 캐시해서 재사용한다.

어떻게:
  이미 지오코딩에 쓰고 있는 카카오 주소검색 API(v2/local/search/address.json)의 응답에
  지번 정보가 같이 들어있다(documents[].address). 추가 API 신청 없이 그대로 쓸 수 있다.
    address.b_code          법정동코드 10자리 -> [0:5]=시군구코드, [5:10]=법정동코드
    address.mountain_yn     산 여부 Y/N      -> platGbCd (N=0 대지, Y=1 산)
    address.main_address_no 본번            -> bun (4자리 0패딩)
    address.sub_address_no  부번            -> ji  (4자리 0패딩)

  도로명주소가 너무 새로워서 address(지번)가 비어있는 경우가 있어, 그때는 이미 캐시에
  있는 위경도로 좌표->주소 변환(v2/local/geo/coord2address.json)을 한 번 더 시도한다.

- geocode.py와 같은 방식으로 하루 한도만큼만 처리하고 중단하며, 이미 처리한 주소는
  pipeline/cache/jibun_cache.json 에 저장해 다음 실행 때 이어서 진행한다.
- 실행: python pipeline/resolve_jibun.py
"""
import json
import os
import time

import requests

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE_DIR = os.path.join(BASE_DIR, "pipeline", "cache")
GEOCODE_CACHE_FILE = os.path.join(CACHE_DIR, "geocode_cache.json")
CACHE_FILE = os.path.join(CACHE_DIR, "jibun_cache.json")
FAILED_FILE = os.path.join(CACHE_DIR, "jibun_failed.json")
CONFIG_FILE = os.path.join(BASE_DIR, "config.local.json")
ADDRESS_SOURCE = os.path.join(BASE_DIR, "국가DB_SID_주소.xlsx")

DAILY_LIMIT = 90000  # 카카오 무료 한도(10만/일)보다 여유를 둔 값 - geocode.py와 동일
SLEEP_SEC = 0.08
SAVE_EVERY = 500

ADDRESS_SEARCH_URL = "https://dapi.kakao.com/v2/local/search/address.json"
COORD_TO_ADDRESS_URL = "https://dapi.kakao.com/v2/local/geo/coord2address.json"


def load_config():
    if not os.path.exists(CONFIG_FILE):
        raise SystemExit(
            f"{CONFIG_FILE} 이 없습니다. config.local.json.example을 복사해서 "
            "config.local.json을 만들고 REST API 키를 넣어주세요."
        )
    with open(CONFIG_FILE, encoding="utf-8") as f:
        return json.load(f)


def load_json(path, default):
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return default


def save_json(path, data):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
    os.replace(tmp, path)


def load_target_addresses(geocode_cache):
    """대상 주소 목록. 원본 주소 엑셀이 있으면 그걸 쓰고(전체), 없으면 이미 지오코딩된
    주소(=지도에 실제로 올라가 있는 주소)만이라도 처리할 수 있게 한다."""
    if os.path.exists(ADDRESS_SOURCE):
        import pandas as pd

        df = pd.read_excel(ADDRESS_SOURCE)
        return sorted(df["ADDRESS1"].dropna().astype(str).str.strip().unique().tolist())
    print(f"경고: {ADDRESS_SOURCE} 없음 - 지오코딩 캐시에 있는 주소만 대상으로 진행합니다.")
    return sorted(geocode_cache.keys())


def build_key(addr_obj, source):
    """카카오 응답의 address(지번) 객체 -> 건축물대장 조회키."""
    b_code = (addr_obj.get("b_code") or "").strip()
    main_no = (addr_obj.get("main_address_no") or "").strip()
    if len(b_code) != 10 or not main_no.isdigit():
        return None
    sub_no = (addr_obj.get("sub_address_no") or "").strip()
    return {
        "sigunguCd": b_code[:5],
        "bjdongCd": b_code[5:],
        "platGbCd": "1" if addr_obj.get("mountain_yn") == "Y" else "0",
        "bun": main_no.zfill(4),
        "ji": (sub_no if sub_no.isdigit() else "0").zfill(4),
        "jibunAddress": addr_obj.get("address_name"),
        "source": source,
    }


def resolve_one(address, rest_key, coord):
    """1차: 도로명주소 검색으로 지번 추출. 2차: 이미 알고 있는 좌표로 좌표->지번 변환."""
    headers = {"Authorization": f"KakaoAK {rest_key}"}

    resp = requests.get(
        ADDRESS_SEARCH_URL, headers=headers, params={"query": address}, timeout=5
    )
    if resp.status_code == 200:
        for doc in resp.json().get("documents", []):
            key = build_key(doc.get("address") or {}, "address_search")
            if key:
                return key

    if not coord:
        return None
    time.sleep(SLEEP_SEC)
    resp = requests.get(
        COORD_TO_ADDRESS_URL,
        headers=headers,
        params={"x": coord["lng"], "y": coord["lat"]},
        timeout=5,
    )
    if resp.status_code != 200:
        return None
    for doc in resp.json().get("documents", []):
        key = build_key(doc.get("address") or {}, "coord2address")
        if key:
            return key
    return None


def seed_from_geocode_cache(cache, geocode_cache):
    """geocode.py가 지번 정보까지 같이 저장해둔 주소는 API를 다시 부르지 않고 그대로 가져온다."""
    seeded = 0
    for addr, info in geocode_cache.items():
        if addr in cache or not isinstance(info, dict):
            continue
        key = build_key(info.get("jibun") or {}, "geocode_cache")
        if key:
            cache[addr] = key
            seeded += 1
    return seeded


def main():
    os.makedirs(CACHE_DIR, exist_ok=True)
    config = load_config()
    rest_key = config["kakao_rest_api_key"]

    geocode_cache = load_json(GEOCODE_CACHE_FILE, {})
    cache = load_json(CACHE_FILE, {})
    failed = load_json(FAILED_FILE, {})

    seeded = seed_from_geocode_cache(cache, geocode_cache)
    if seeded:
        print(f"지오코딩 캐시에서 그대로 가져온 주소: {seeded}건 (API 호출 없음)")
        save_json(CACHE_FILE, cache)

    addresses = load_target_addresses(geocode_cache)
    print(f"대상 주소: {len(addresses)}건 / 이미 변환: {len(cache)}건 / 실패기록: {len(failed)}건")

    todo = [a for a in addresses if a not in cache and a not in failed]
    if not todo:
        print("모든 주소의 건축물대장 조회키 변환 완료.")
        return

    batch = todo[:DAILY_LIMIT]
    print(f"이번 실행 처리 예정: {len(batch)}건 (일일 한도 {DAILY_LIMIT}건 기준)")

    done = 0
    for addr in batch:
        try:
            result = resolve_one(addr, rest_key, geocode_cache.get(addr))
        except requests.RequestException as e:
            print(f"[네트워크 오류] {addr}: {e}")
            continue
        if result:
            cache[addr] = result
        else:
            failed[addr] = True
        done += 1
        if done % SAVE_EVERY == 0:
            save_json(CACHE_FILE, cache)
            save_json(FAILED_FILE, failed)
            print(f"  진행 {done}/{len(batch)} (누적 성공 {len(cache)}건, 실패 {len(failed)}건)")
        time.sleep(SLEEP_SEC)

    save_json(CACHE_FILE, cache)
    save_json(FAILED_FILE, failed)

    unique_pnu = {
        (v["sigunguCd"], v["bjdongCd"], v["platGbCd"], v["bun"], v["ji"]) for v in cache.values()
    }
    remaining = len(addresses) - len(cache) - len(failed)
    print(f"\n이번 실행 처리: {done}건")
    print(f"누적 변환 성공: {len(cache)}건 / 실패: {len(failed)}건")
    print(f"고유 지번(건축물대장 조회 대상): {len(unique_pnu)}건")
    if remaining > 0:
        print(f"남은 주소: {remaining}건 -> 내일 다시 실행하면 이어서 진행됩니다.")
    else:
        print("모든 주소 변환 완료 -> 다음 단계: python pipeline/fetch_building_ledger.py")


if __name__ == "__main__":
    main()
