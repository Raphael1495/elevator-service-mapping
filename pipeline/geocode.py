"""
주소 -> 위경도 지오코딩 (카카오 로컬 API)

- 국가DB_SID_주소.xlsx 의 고유 주소를 대상으로 카카오 주소검색 API를 호출한다.
- 하루 무료 한도(10만 건)를 넘지 않도록 DAILY_LIMIT 만큼만 처리하고 중단한다.
- 이미 처리한 주소는 pipeline/cache/geocode_cache.json 에 저장해두고,
  다음 실행 때는 남은 주소부터 이어서 처리한다 (하루 한 번씩 며칠에 걸쳐 실행).
- 실행: python pipeline/geocode.py
"""
import json
import os
import time

import pandas as pd
import requests

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE_DIR = os.path.join(BASE_DIR, "pipeline", "cache")
CACHE_FILE = os.path.join(CACHE_DIR, "geocode_cache.json")
FAILED_FILE = os.path.join(CACHE_DIR, "geocode_failed.json")
CONFIG_FILE = os.path.join(BASE_DIR, "config.local.json")
ADDRESS_SOURCE = os.path.join(BASE_DIR, "국가DB_SID_주소.xlsx")

DAILY_LIMIT = 90000  # 카카오 무료 한도(10만/일)보다 여유를 둔 값
SLEEP_SEC = 0.08  # 초당 약 12건 - 카카오 QPS 제한 내 안전 마진
SAVE_EVERY = 500


def load_config():
    if not os.path.exists(CONFIG_FILE):
        raise SystemExit(
            f"{CONFIG_FILE} 이 없습니다. config.local.json.example을 복사해서 "
            "config.local.json을 만들고 REST API 키를 넣어주세요."
        )
    with open(CONFIG_FILE, encoding="utf-8") as f:
        return json.load(f)


def load_unique_addresses():
    df = pd.read_excel(ADDRESS_SOURCE)
    return sorted(df["ADDRESS1"].dropna().astype(str).str.strip().unique().tolist())


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


def geocode_one(address, rest_key):
    url = "https://dapi.kakao.com/v2/local/search/address.json"
    headers = {"Authorization": f"KakaoAK {rest_key}"}
    resp = requests.get(url, headers=headers, params={"query": address}, timeout=5)
    if resp.status_code != 200:
        return None
    docs = resp.json().get("documents", [])
    if not docs:
        return None
    d = docs[0]
    return {"lat": float(d["y"]), "lng": float(d["x"])}


def main():
    os.makedirs(CACHE_DIR, exist_ok=True)
    config = load_config()
    rest_key = config["kakao_rest_api_key"]

    print("고유 주소 목록 로딩 중...")
    addresses = load_unique_addresses()
    print(f"전체 고유 주소: {len(addresses)}건")

    cache = load_json(CACHE_FILE, {})
    failed = load_json(FAILED_FILE, {})

    todo = [a for a in addresses if a not in cache and a not in failed]
    print(f"이미 완료: {len(cache)}건 / 실패기록: {len(failed)}건 / 남은 작업: {len(todo)}건")

    if not todo:
        print("모든 주소 지오코딩 완료.")
        return

    batch = todo[:DAILY_LIMIT]
    print(f"이번 실행 처리 예정: {len(batch)}건 (일일 한도 {DAILY_LIMIT}건 기준)")

    done_count = 0
    for addr in batch:
        try:
            result = geocode_one(addr, rest_key)
        except requests.RequestException as e:
            print(f"[네트워크 오류] {addr}: {e}")
            continue
        if result:
            cache[addr] = result
        else:
            failed[addr] = True
        done_count += 1
        if done_count % SAVE_EVERY == 0:
            save_json(CACHE_FILE, cache)
            save_json(FAILED_FILE, failed)
            print(f"  진행 {done_count}/{len(batch)} (누적 성공 {len(cache)}건, 실패 {len(failed)}건)")
        time.sleep(SLEEP_SEC)

    save_json(CACHE_FILE, cache)
    save_json(FAILED_FILE, failed)

    remaining = len(addresses) - len(cache) - len(failed)
    print(f"\n이번 실행 처리: {done_count}건")
    print(f"누적 지오코딩 성공: {len(cache)}건 / 실패(주소 매칭 안됨): {len(failed)}건")
    if remaining > 0:
        print(f"남은 주소: {remaining}건 -> 내일 이 스크립트를 다시 실행하면 이어서 진행됩니다.")
    else:
        print("모든 주소 지오코딩 완료.")


if __name__ == "__main__":
    main()
