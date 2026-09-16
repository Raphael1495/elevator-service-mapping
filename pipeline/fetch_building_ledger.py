"""
건축물대장 표제부 수집 (공공데이터포털 - 국토교통부_건축HUB_건축물대장정보 서비스).

- 조회 단위는 "지번"이다. resolve_jibun.py가 만든 pipeline/cache/jibun_cache.json 의
  시군구코드/법정동코드/대지구분/본번/부번을 그대로 요청 파라미터로 넣는다.
  주소 여러 개가 같은 지번을 가리키면 호출은 한 번만 한다(아파트 단지 등에서 크게 절약됨).
- 한 지번에 동이 여러 개면 표제부도 여러 건 내려온다. 여기서는 그 목록을 가공하지 않고
  전부 캐시에 저장하고, "어느 동이 우리 현장인가" 판단은 현장명(BULDNM)을 들고 있는
  build_dataset.py 에서 building_ledger.match_ledger_row()로 한다.
- geocode.py와 같이 하루 한도만큼만 처리하고 중단하며, 결과는
  pipeline/cache/building_ledger.json 에 쌓아 다음 실행 때 이어서 진행한다.

실행:
  python pipeline/fetch_building_ledger.py --self-test    # 인증키/응답 확인 (1건만 호출)
  python pipeline/fetch_building_ledger.py                # 전체 수집 (하루 한도만큼)
  python pipeline/fetch_building_ledger.py --region "서울특별시 강남구"   # 특정 지역만
  python pipeline/fetch_building_ledger.py --limit 500    # 이번 실행 호출 수 직접 지정
"""
import argparse
import json
import os
import sys
import time
from urllib.parse import unquote

import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from building_ledger import LEDGER_FIELDS, pnu_key  # noqa: E402

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE_DIR = os.path.join(BASE_DIR, "pipeline", "cache")
JIBUN_CACHE_FILE = os.path.join(CACHE_DIR, "jibun_cache.json")
LEDGER_CACHE_FILE = os.path.join(CACHE_DIR, "building_ledger.json")
LEDGER_FAILED_FILE = os.path.join(CACHE_DIR, "building_ledger_failed.json")
CONFIG_FILE = os.path.join(BASE_DIR, "config.local.json")

API_URL = "https://apis.data.go.kr/1613000/BldRgstHubService/getBrTitleInfo"

# 공공데이터포털 트래픽 제한: 개발계정은 보통 일 1,000건, 운영계정 승인 후 상향(흔히 10만/일).
# 지금 값은 운영계정 기준이며, 개발계정 단계에서는 --limit 900 정도로 낮춰서 실행한다.
DAILY_LIMIT = 90000
SLEEP_SEC = 0.05
SAVE_EVERY = 200
NUM_OF_ROWS = 100  # 한 지번의 동 수 - 대단지도 100동을 넘는 경우는 사실상 없음
MAX_RETRY = 3

# 공공데이터포털이 정상 처리로 내려주는 결과코드. "00"이 정상이고,
# 자료가 없는 경우도 정상 응답(totalCount=0)으로 온다.
RESULT_CODE_OK = {"00", "0"}
# 재시도해도 소용없는(키/신청 문제) 오류 - 만나면 즉시 중단하고 사람이 조치해야 한다.
FATAL_RESULT_CODES = {
    "30": "등록되지 않은 서비스키입니다 (일반 인증키 확인 필요)",
    "31": "활용기간이 만료된 인증키입니다",
    "22": "일일 트래픽 한도를 초과했습니다 (내일 다시 실행하거나 운영계정 증량 신청)",
    "20": "서비스 접근이 거부되었습니다 (활용 신청 승인 상태 확인)",
}


def normalize_service_key(key):
    """공공데이터포털 마이페이지는 일반 인증키를 Encoding / Decoding 두 가지로 보여준다.
    requests가 파라미터를 알아서 URL 인코딩하므로 여기서는 Decoding 값이 필요한데,
    Encoding 값(%2B, %3D 등이 섞인 문자열)을 붙여넣는 실수가 흔해서 자동으로 되돌린다."""
    if "%" in key:
        decoded = unquote(key)
        if decoded != key:
            print("인증키가 Encoding 값으로 보여서 Decoding 값으로 변환해서 사용합니다.")
            return decoded
    return key


def load_config():
    if not os.path.exists(CONFIG_FILE):
        raise SystemExit(
            f"{CONFIG_FILE} 이 없습니다. config.local.json.example을 복사해서 "
            "config.local.json을 만들고 인증키를 넣어주세요."
        )
    with open(CONFIG_FILE, encoding="utf-8") as f:
        config = json.load(f)
    key = config.get("data_go_kr_service_key")
    if not key:
        raise SystemExit(
            "config.local.json 에 data_go_kr_service_key 가 없습니다. "
            "공공데이터포털 마이페이지의 '일반 인증키(Decoding)' 값을 넣어주세요."
        )
    config["data_go_kr_service_key"] = normalize_service_key(key.strip())
    return config


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


class FatalApiError(Exception):
    """재시도해도 해결되지 않는 오류 (인증키/트래픽/승인 상태)."""


def _extract_items(payload):
    """공공데이터포털 응답에서 item 목록을 꺼낸다.
    - 자료 없음이면 items가 빈 문자열/None 으로 온다.
    - 1건이면 item이 리스트가 아니라 단일 객체로 온다."""
    body = (payload.get("response") or {}).get("body") or {}
    items = body.get("items")
    if not items:
        return []
    item = items.get("item") if isinstance(items, dict) else items
    if item is None:
        return []
    return item if isinstance(item, list) else [item]


def _check_header(payload, raw_text):
    header = (payload.get("response") or {}).get("header") or {}
    code = str(header.get("resultCode", "")).strip()
    if code in RESULT_CODE_OK:
        return
    if code in FATAL_RESULT_CODES:
        raise FatalApiError(f"resultCode={code} - {FATAL_RESULT_CODES[code]}")
    raise requests.RequestException(
        f"resultCode={code} msg={header.get('resultMsg')} raw={raw_text[:200]}"
    )


def _parse_response(resp):
    """정상일 때는 JSON, 인증키 오류 등에서는 XML(OpenAPI_ServiceResponse)이 오는 경우가 있다."""
    text = resp.text
    try:
        payload = resp.json()
    except ValueError:
        if "SERVICE_KEY_IS_NOT_REGISTERED" in text or "<returnReasonCode>30" in text:
            raise FatalApiError(
                "등록되지 않은 서비스키입니다 - 공공데이터포털의 '일반 인증키(Decoding)' 값인지, "
                "해당 API 활용신청이 승인됐는지 확인해주세요."
            )
        raise requests.RequestException(f"JSON 파싱 실패 (status={resp.status_code}) raw={text[:200]}")
    _check_header(payload, text)
    return _extract_items(payload)


def _request(jibun, service_key, plat_gb_cd):
    params = {
        "serviceKey": service_key,
        "sigunguCd": jibun["sigunguCd"],
        "bjdongCd": jibun["bjdongCd"],
        "platGbCd": plat_gb_cd,
        "bun": jibun["bun"],
        "ji": jibun["ji"],
        "numOfRows": NUM_OF_ROWS,
        "pageNo": 1,
        "_type": "json",
    }
    last_error = None
    for attempt in range(MAX_RETRY):
        try:
            resp = requests.get(API_URL, params=params, timeout=10)
            items = _parse_response(resp)
            return [{f: item.get(f) for f in LEDGER_FIELDS if item.get(f) not in (None, "")} for item in items]
        except FatalApiError:
            raise
        except requests.RequestException as e:
            last_error = e
            time.sleep(2 ** attempt)
    raise requests.RequestException(str(last_error))


def fetch_one(jibun, service_key):
    """지번 하나의 표제부 목록을 가져온다. 우리가 쓰는 필드만 남겨서 반환.

    대지구분(platGbCd)은 카카오 응답의 산 여부로 0/1을 정하는데, 실제 대장이 다른
    구분(예: 블록)으로 등록돼 있으면 0건이 나온다. 0건일 때는 대지구분을 비우고
    (= 전체) 한 번 더 조회해서 놓치는 건을 줄인다."""
    rows = _request(jibun, service_key, jibun["platGbCd"])
    if rows:
        return rows
    return _request(jibun, service_key, "")


def build_targets(jibun_cache, region_filter):
    """지번 -> 조회 파라미터. 주소 여러 개가 같은 지번이면 한 번만 호출하도록 묶는다.
    --region 을 준 경우에만 build_dataset과 같은 기준으로 지역을 판정해서 걸러낸다."""
    region_of = None
    if region_filter:
        from build_dataset import region_of  # 지역 판정 기준을 한 곳(build_dataset)에만 두기 위함

    targets = {}
    for address, jibun in jibun_cache.items():
        if region_of and region_of(address) not in region_filter:
            continue
        targets.setdefault(pnu_key(jibun), jibun)
    return targets


def self_test(service_key):
    """인증키와 응답 형식만 확인하는 1건 호출 (서울 강남구 역삼동 736-17)."""
    sample = {"sigunguCd": "11680", "bjdongCd": "10100", "platGbCd": "0", "bun": "0736", "ji": "0017"}
    print(f"테스트 호출: {API_URL}")
    print(f"파라미터: {sample}")
    rows = fetch_one(sample, service_key)
    print(f"\n표제부 {len(rows)}건 수신")
    for row in rows[:3]:
        print(json.dumps(row, ensure_ascii=False, indent=2))
    if rows:
        print("\n인증키 정상 동작 - 전체 수집을 진행할 수 있습니다.")
    else:
        print("\n응답은 정상인데 자료가 0건입니다(그 지번에 건축물대장이 없을 수 있음).")
        print("다른 지번으로도 0건이면 활용신청 승인 상태를 확인해주세요.")


def main():
    parser = argparse.ArgumentParser(description="건축물대장 표제부 수집")
    parser.add_argument("--self-test", action="store_true", help="인증키 확인용 1건 호출")
    parser.add_argument("--limit", type=int, default=DAILY_LIMIT, help=f"이번 실행 최대 호출 수 (기본 {DAILY_LIMIT})")
    parser.add_argument("--region", action="append", help='특정 지역만 수집 (예: --region "서울특별시 강남구")')
    args = parser.parse_args()

    os.makedirs(CACHE_DIR, exist_ok=True)
    service_key = load_config()["data_go_kr_service_key"]

    if args.self_test:
        self_test(service_key)
        return

    jibun_cache = load_json(JIBUN_CACHE_FILE, {})
    if not jibun_cache:
        raise SystemExit(
            f"{JIBUN_CACHE_FILE} 이 비어있습니다. 먼저 python pipeline/resolve_jibun.py 를 실행하세요."
        )

    ledger = load_json(LEDGER_CACHE_FILE, {})
    failed = load_json(LEDGER_FAILED_FILE, {})

    region_filter = set(args.region) if args.region else None
    targets = build_targets(jibun_cache, region_filter)
    scope = f"({', '.join(sorted(region_filter))})" if region_filter else "(전국)"
    print(f"주소 {len(jibun_cache)}건 -> 고유 지번 {len(targets)}건 {scope}")
    print(f"이미 수집: {len(ledger)}건 / 실패기록: {len(failed)}건")

    todo = [(k, v) for k, v in targets.items() if k not in ledger and k not in failed]
    if not todo:
        print("이 범위의 건축물대장 수집 완료 -> 다음 단계: python pipeline/build_dataset.py")
        return

    batch = todo[: args.limit]
    print(f"이번 실행 처리 예정: {len(batch)}건 (남은 작업 {len(todo)}건)")

    done = 0
    empty = 0
    try:
        for key, jibun in batch:
            try:
                rows = fetch_one(jibun, service_key)
            except FatalApiError as e:
                print(f"\n[중단] {e}")
                break
            except requests.RequestException as e:
                failed[key] = str(e)[:200]
                print(f"[실패] {key}: {e}")
                done += 1
                continue
            if rows:
                ledger[key] = rows
            else:
                # 자료 없음도 결과로 기록해서 다음 실행 때 다시 부르지 않게 한다.
                ledger[key] = []
                empty += 1
            done += 1
            if done % SAVE_EVERY == 0:
                save_json(LEDGER_CACHE_FILE, ledger)
                save_json(LEDGER_FAILED_FILE, failed)
                print(f"  진행 {done}/{len(batch)} (누적 {len(ledger)}건, 자료없음 {empty}건, 실패 {len(failed)}건)")
            time.sleep(SLEEP_SEC)
    except KeyboardInterrupt:
        print("\n사용자 중단 - 여기까지 진행분을 저장합니다.")

    save_json(LEDGER_CACHE_FILE, ledger)
    save_json(LEDGER_FAILED_FILE, failed)

    multi_dong = sum(1 for rows in ledger.values() if len(rows) > 1)
    remaining = len(targets) - len(ledger) - len(failed)
    print(f"\n이번 실행 처리: {done}건 (자료없음 {empty}건)")
    print(f"누적 수집 지번: {len(ledger)}건 / 실패: {len(failed)}건")
    print(f"동이 2개 이상인 지번(현장명 매칭 필요): {multi_dong}건")
    if remaining > 0:
        print(f"남은 지번: {remaining}건 -> 내일 다시 실행하면 이어서 진행됩니다.")
    else:
        print("수집 완료 -> 다음 단계: python pipeline/build_dataset.py")


if __name__ == "__main__":
    main()
