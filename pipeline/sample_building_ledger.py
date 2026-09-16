"""
건축물대장 API가 실제로 어떤 값을 주는지 확인하기 위한 소규모 샘플 스크립트.

이 저장소에 이미 있는 실제 승강기 현장 8곳(건물 유형이 서로 다르게 골라둠)을 대상으로
  1) 카카오 주소검색으로 도로명주소 -> 지번 변환
  2) 건축물대장 표제부(getBrTitleInfo) 조회
를 실행해서 결과를 엑셀 파일로 저장한다.

build_dataset.py가 쓰는 pipeline/building_ledger.py의 LEDGER_FIELDS는 필요한 것만
추려서 캐시 용량을 줄인 목록인 반면, 이 스크립트는 "API가 뭘 주는지 전부 보고 싶다"는
목적이므로 응답 필드를 트리밍하지 않고 있는 그대로 전부 저장한다.

전제: apis.data.go.kr / dapi.kakao.com 에 접속 가능한 환경에서 실행해야 한다.
      (Claude Code 클라우드 세션은 조직 정책으로 이 도메인들이 막혀있어 여기서는 못 돌림)

실행: python pipeline/sample_building_ledger.py
결과: pipeline/sample_building_ledger_output.xlsx
       - "조회결과" 시트: API 응답 원본 필드 전부 (동 여러 개면 여러 행)
       - "요약" 시트: 우리가 원한 값(주용도/기타용도 등)만 8곳 기준으로 한눈에
"""
import json
import os
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fetch_building_ledger import (  # noqa: E402
    API_URL,
    FatalApiError,
    _parse_response,
    normalize_service_key,
)
from resolve_jibun import ADDRESS_SEARCH_URL, build_key  # noqa: E402

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_FILE = os.path.join(BASE_DIR, "config.local.json")
OUTPUT_FILE = os.path.join(BASE_DIR, "pipeline", "sample_building_ledger_output.xlsx")

# 이 저장소에 이미 커밋된 실제 현장 중, 국가DB 설치신고 용도(BULDPRPOS)가 서로 다르게
# 8곳을 골랐다 (data/regions + data/details 에서 추출 - 전국 region 파일을 무작위로
# 훑어서 8개 대분류가 하나씩 나오는 첫 매치를 채택. 매번 다른 현장을 보고 싶으면
# 이 리스트를 다른 주소로 바꿔서 실행해도 된다).
SAMPLE_SITES = [
    {"region": "세종특별자치시 달빛2로", "address": "세종특별자치시 달빛2로 29",
     "name": "아름유치원", "elevatorNo": "5015456", "installReportedType": "교육연구시설-유치원"},
    {"region": "세종특별자치시 한솔동로", "address": "세종특별자치시 한솔동로 26",
     "name": "2생활권 환승주차장", "elevatorNo": "5067867", "installReportedType": "업무시설-기타"},
    {"region": "세종특별자치시 장군면", "address": "세종특별자치시 장군면 봉안길 70",
     "name": "세종스파", "elevatorNo": "5077647", "installReportedType": "제1종근린생활시설-기타"},
    {"region": "세종특별자치시 다솜로", "address": "세종특별자치시 다솜로 7 (어진동)",
     "name": "한뜰마을5단지아파트", "elevatorNo": "5091839", "installReportedType": "공동주택-분양"},
    {"region": "경상북도 청도군", "address": "경상북도 청도군 화양읍 고수구길 2",
     "name": "선원빌딩", "elevatorNo": "6037676", "installReportedType": "제2종근린생활시설-기타"},
    {"region": "강원특별자치도 화천군", "address": "강원특별자치도 화천군 화천읍 강변로 111",
     "name": "화천군보건의료원", "elevatorNo": "4004455", "installReportedType": "의료시설-기타"},
    {"region": "충청남도 부여군", "address": "충청남도 부여군 규암면 서궁로 129",
     "name": "부여백제파크아트빌라스", "elevatorNo": "5079679", "installReportedType": "숙박시설-숙박시설"},
    {"region": "충청남도 계룡시", "address": "충청남도 계룡시 계룡대로 304",
     "name": "홈플러스계룡점", "elevatorNo": "5029957", "installReportedType": "판매시설-기타"},
]

SUMMARY_FIELDS = [
    "dongNm", "bldNm", "mainPurpsCdNm", "etcPurps", "mainAtchGbCdNm",
    "useAprDay", "totArea", "grndFlrCnt", "ugrndFlrCnt", "hhldCnt",
    "rideUseElvtCnt", "emgenUseElvtCnt",
]


def load_config():
    if not os.path.exists(CONFIG_FILE):
        raise SystemExit(
            f"{CONFIG_FILE} 이 없습니다. config.local.json.example을 복사해서 "
            "config.local.json을 만들고 kakao_rest_api_key / data_go_kr_service_key를 넣어주세요."
        )
    with open(CONFIG_FILE, encoding="utf-8") as f:
        config = json.load(f)
    if not config.get("kakao_rest_api_key") or not config.get("data_go_kr_service_key"):
        raise SystemExit("config.local.json 에 kakao_rest_api_key / data_go_kr_service_key 가 모두 필요합니다.")
    config["data_go_kr_service_key"] = normalize_service_key(config["data_go_kr_service_key"].strip())
    return config


def resolve_jibun(address, kakao_key):
    """resolve_jibun.py의 로직을 그대로 쓰되, 좌표 캐시가 없어도 되도록 여기서 자체 처리."""
    headers = {"Authorization": f"KakaoAK {kakao_key}"}
    resp = requests.get(ADDRESS_SEARCH_URL, headers=headers, params={"query": address}, timeout=5)
    if resp.status_code == 200:
        for doc in resp.json().get("documents", []):
            key = build_key(doc.get("address") or {}, "address_search")
            if key:
                return key
    return None


def fetch_raw(jibun, service_key):
    """트리밍 없이 표제부 응답을 있는 그대로 반환. platGbCd 0건이면 빈 값으로 한 번 더."""
    for plat_gb_cd in (jibun["platGbCd"], ""):
        params = {
            "serviceKey": service_key,
            "sigunguCd": jibun["sigunguCd"],
            "bjdongCd": jibun["bjdongCd"],
            "platGbCd": plat_gb_cd,
            "bun": jibun["bun"],
            "ji": jibun["ji"],
            "numOfRows": 100,
            "pageNo": 1,
            "_type": "json",
        }
        resp = requests.get(API_URL, params=params, timeout=10)
        items = _parse_response(resp)
        if items:
            return items
    return []


def main():
    config = load_config()
    kakao_key = config["kakao_rest_api_key"]
    service_key = config["data_go_kr_service_key"]

    all_rows = []
    summary_rows = []
    print(f"샘플 {len(SAMPLE_SITES)}곳 조회 시작\n")

    for site in SAMPLE_SITES:
        addr = site["address"]
        print(f"[{site['name']}] {addr}")
        jibun = resolve_jibun(addr, kakao_key)
        if not jibun:
            print("  -> 지번 변환 실패 (카카오 주소검색에서 못 찾음)")
            summary_rows.append({**site, "결과": "지번 변환 실패"})
            continue
        print(f"  지번: {jibun['jibunAddress']} (sigunguCd={jibun['sigunguCd']}, bjdongCd={jibun['bjdongCd']}, "
              f"bun={jibun['bun']}, ji={jibun['ji']})")

        try:
            items = fetch_raw(jibun, service_key)
        except FatalApiError as e:
            raise SystemExit(f"\n[중단] {e}")

        if not items:
            print("  -> 건축물대장 조회 결과 0건")
            summary_rows.append({**site, "결과": "대장 없음(0건)"})
            continue

        print(f"  -> 표제부 {len(items)}건 (동 {len(items)}개)")
        for item in items:
            all_rows.append({"조회_현장명": site["name"], "조회_주소": addr, **item})

        # 요약 시트용 - 동이 여러 개면 국가DB 현장명과 가장 비슷한 것 하나만 대표로.
        # (여기서는 정밀 매칭 대신 그냥 첫 번째 값을 보여준다 - 실제 파이프라인은
        # building_ledger.match_ledger_row()가 이 판단을 한다)
        rep = items[0]
        summary_rows.append({
            **site,
            "결과": f"{len(items)}건 중 1번째 표시" if len(items) > 1 else "1건",
            **{f: rep.get(f) for f in SUMMARY_FIELDS},
        })

    if not all_rows:
        raise SystemExit("\n조회된 건축물대장이 하나도 없습니다. 인증키/네트워크를 확인해주세요.")

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="요약", index=False)
        pd.json_normalize(all_rows).to_excel(writer, sheet_name="조회결과(원본필드 전체)", index=False)

    print(f"\n완료: {OUTPUT_FILE}")
    print(f"  요약 {len(summary_rows)}행 / 원본 응답 {len(all_rows)}행")


if __name__ == "__main__":
    main()
