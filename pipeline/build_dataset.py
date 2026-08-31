"""
국가DB(aa 시트) + 주소 매핑 + 지오코딩 캐시 + 현대엘리베이터 자체 관리대수(원본, 프로젝트번호)를
합쳐서 지도에서 바로 쓸 수 있는 데이터를 만든다.

- 전국 데이터를 파일 하나로 만들면 89만 건 기준 수백MB가 되어 GitHub 파일 용량 제한(100MB)에
  걸리고 브라우저도 느려지므로, 주소 앞자리(시/도)로 나눠서 data/regions/{시도명}.json 으로 저장한다.
- data/regions_manifest.json 에는 지역별 건수와 위경도 범위(bbox)를 담아, 프론트에서
  현재 지도 화면에 걸치는 지역 파일만 불러올 수 있게 한다.
- 아직 지오코딩이 끝나지 않은 주소는 건너뛴다 (geocode.py를 며칠에 걸쳐 실행하면서
  이 스크립트를 다시 돌리면 점점 더 많은 데이터가 채워진다).
- 프로젝트번호(projectNo)는 현대엘리베이터가 직접 관리하는 건에만 존재한다 (없으면 null).
- 실행: python pipeline/build_dataset.py [건수제한]
  건수제한을 생략하면 지오코딩된 것 전체를 반영한다.
"""
import json
import os
import sys
from collections import Counter, defaultdict
from datetime import date

import pandas as pd
import pyxlsb

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE_FILE = os.path.join(BASE_DIR, "pipeline", "cache", "geocode_cache.json")
ADDRESS_SOURCE = os.path.join(BASE_DIR, "국가DB_SID_주소.xlsx")
DB_SOURCE = os.path.join(BASE_DIR, "국가DB_260805.xlsb")
DB_SHEET = "aa"
MGMT_SOURCE = os.path.join(BASE_DIR, "2026.07월 관리대수(원본).xlsb")
MGMT_SHEET = "download"
REGIONS_DIR = os.path.join(BASE_DIR, "data", "regions")
MANIFEST_FILE = os.path.join(BASE_DIR, "data", "regions_manifest.json")
SEARCH_INDEX_FILE = os.path.join(BASE_DIR, "data", "search_index.json")

NEEDED_COLUMNS = [
    "ELEVATORNO",
    "BULDNM",
    "MANUFACTURERNAME",
    "MNTCPNYNM",
    "ELVTRSTTS",
    "ELVTRKINDNM",
    "ELVTRMODEL",
    "INSTALLATIONDE",
    "FRSTINSTALLATIONDE",
]


def normalize_key(value):
    """엘리베이터 번호를 파일마다 다른 표기(문자열 0패딩 / float)에서 공통 정수 문자열로 통일."""
    if value is None:
        return None
    try:
        return str(int(float(value)))
    except (TypeError, ValueError):
        return str(value).strip()


def pad_elevator_no(eno):
    """국가DB 공식 표기(7자리, 0패딩)로 복원."""
    return eno.zfill(7) if eno else eno


def pad_project_no(pjt):
    """프로젝트번호(6자리, 0패딩)로 복원."""
    return pjt.zfill(6) if pjt else pjt


# 원본 주소에 개칭 전/후 표기가 섞여 있어(예: 강원도 -> 강원특별자치도, 2023년 개칭)
# 같은 지역이 여러 키로 쪼개지지 않도록 시/도명을 최신 공식 명칭으로 통일한다.
SIDO_ALIASES = {
    "강원도": "강원특별자치도",
    "강원": "강원특별자치도",
    "전라북도": "전북특별자치도",
    "전북": "전북특별자치도",
    "경기": "경기도",
    "서울": "서울특별시",
    "부산": "부산광역시",
    "대구": "대구광역시",
    "인천": "인천광역시",
    "광주": "광주광역시",
    "대전": "대전광역시",
    "울산": "울산광역시",
    "세종": "세종특별자치시",
    "충북": "충청북도",
    "충남": "충청남도",
    "전남": "전라남도",
    "경북": "경상북도",
    "경남": "경상남도",
    "제주": "제주특별자치도",
}

KNOWN_SIDO = set(SIDO_ALIASES.values()) | {
    "서울특별시", "부산광역시", "대구광역시", "인천광역시", "광주광역시", "대전광역시",
    "울산광역시", "세종특별자치시", "경기도", "충청북도", "충청남도", "전라남도",
    "경상북도", "경상남도", "제주특별자치도", "강원특별자치도", "전북특별자치도",
}

# 일반시(광역시가 아닌 시) 산하에 구가 있는 도시들 - 주소에 중간의 시 이름이
# 빠진 채로("경기도 일산동구 ..." 처럼) 들어오면 그 구 하나만 별도의 엉뚱한
# 지역 버킷으로 떨어져 나가므로, 구 이름만 보고 원래 시로 되돌려준다.
GU_TO_CITY = {
    "장안구": "수원시", "권선구": "수원시", "팔달구": "수원시", "영통구": "수원시",
    "수정구": "성남시", "중원구": "성남시", "분당구": "성남시",
    "만안구": "안양시", "동안구": "안양시",
    "상록구": "안산시", "단원구": "안산시",
    "덕양구": "고양시", "일산동구": "고양시", "일산서구": "고양시",
    "처인구": "용인시", "기흥구": "용인시", "수지구": "용인시",
    "상당구": "청주시", "서원구": "청주시", "흥덕구": "청주시", "청원구": "청주시",
    "동남구": "천안시", "서북구": "천안시",
    "완산구": "전주시", "덕진구": "전주시",
    "의창구": "창원시", "성산구": "창원시", "마산합포구": "창원시", "마산회원구": "창원시", "진해구": "창원시",
}

# 시-구 사이 띄어쓰기가 빠져 하나로 붙어버린 표기("부천시원미구" 등) 보정.
# 부천시는 2016년에 구가 폐지되어 지금은 "부천시"만 쓰지만, 옛날식 표기가
# 남아있는 주소가 일부 있어 별도의 엉뚱한 지역으로 떨어져 나갔음.
SIGUNGU_ALIASES = {
    "부천시원미구": "부천시",
    "부천시소사구": "부천시",
    "부천시오정구": "부천시",
}


def region_of(address):
    """주소의 '시/도 + 시/군/구' 두 토큰을 지역 구분 키로 사용 (시/도 하나로만 나누면
    경기도처럼 인구 밀집 지역이 혼자 거대 파일이 되어버리는 문제를 피하기 위함)."""
    if not address:
        return "기타"
    tokens = address.strip().split(" ")
    sido = SIDO_ALIASES.get(tokens[0], tokens[0])
    if sido not in KNOWN_SIDO:
        return "기타"
    sigungu = tokens[1] if len(tokens) > 1 else ""
    sigungu = SIGUNGU_ALIASES.get(sigungu, sigungu)
    sigungu = GU_TO_CITY.get(sigungu, sigungu)
    return (sido + " " + sigungu).strip()


# 화면(index.html)의 classifyCompany()와 동일한 규칙 - 지역 랭킹 패널의 업체별
# 집계가 지도 대시보드 숫자와 어긋나지 않도록 같은 분류 기준을 파이썬에도 둔다.
def classify_company(name):
    if not name:
        return "기타"
    n = str(name).upper()
    if "오티스" in n or "OTIS" in n or "엘지산전" in str(name) or "LG산전" in n:
        return "오티스"
    if "현대" in str(name):
        return "현대"
    if "티센" in str(name) or "티케이" in str(name) or "THYSSEN" in n or "TK" in n:
        return "티케이"
    if "미쓰비시" in str(name) or "MITSUBISHI" in n:
        return "미쓰비시"
    return "기타"


# 승강기안전관리법 기준 정밀안전검사 대상(설치검사일로부터 15년 경과)
OLD_ELEVATOR_YEARS = 15


def is_old(first_install_date):
    if not first_install_date:
        return False
    try:
        y, m, d = str(first_install_date)[:10].split("-")
        installed = date(int(y), int(m), int(d))
    except (ValueError, AttributeError):
        return False
    return (date.today() - installed).days / 365.25 >= OLD_ELEVATOR_YEARS


def count_units_by_address(addr_map):
    """같은 주소에 설치된 승강기번호 개수 - 국가DB 전체 기준(지오코딩 여부와 무관).
    건물명(BULDNM)은 "혜성빌딩"처럼 전국에 흔한 이름이 많아서 서로 다른 건물이
    합쳐지는 문제가 있었음 -> 주소(사실상 유일한 건물 식별자) 기준으로 변경."""
    return Counter(addr_map.values())


def load_cache():
    if not os.path.exists(CACHE_FILE):
        raise SystemExit(f"{CACHE_FILE} 이 없습니다. 먼저 pipeline/geocode.py를 실행하세요.")
    with open(CACHE_FILE, encoding="utf-8") as f:
        return json.load(f)


def load_addr_map():
    df = pd.read_excel(ADDRESS_SOURCE)
    df["ADDRESS1"] = df["ADDRESS1"].astype(str).str.strip()
    return {normalize_key(k): v for k, v in zip(df["ELEVATORNO"], df["ADDRESS1"])}


def load_hyundai_management_map():
    """현대엘리베이터 자체 관리대수(원본)에서 승강기번호 -> {프로젝트번호, 담당팀} 매핑.
    승강기번호가 없는 행(약 34건)은 조인할 수 없으니 그냥 버린다."""
    if not os.path.exists(MGMT_SOURCE):
        print(f"경고: {MGMT_SOURCE} 없음 - 프로젝트번호/담당팀 없이 진행")
        return {}
    result = {}
    with pyxlsb.open_workbook(MGMT_SOURCE) as wb:
        with wb.get_sheet(MGMT_SHEET) as sheet:
            rows = sheet.rows()
            header = [c.v for c in next(rows)]
            idx_eno = header.index("승강기번호")
            idx_pjt = header.index("원PJT")
            idx_team = header.index("팀")
            for row in rows:
                vals = [c.v for c in row]
                eno = normalize_key(vals[idx_eno])
                if not eno:
                    continue
                pjt = vals[idx_pjt]
                team = vals[idx_team]
                result[eno] = {
                    "projectNo": normalize_key(pjt) if pjt is not None else None,
                    "team": team.strip() if isinstance(team, str) and team.strip() else None,
                }
    return result


def main():
    limit = int(sys.argv[1]) if len(sys.argv) > 1 else None

    cache = load_cache()
    addr_map = load_addr_map()
    mgmt_map = load_hyundai_management_map()
    unit_counts = count_units_by_address(addr_map)
    print(f"지오코딩 캐시: {len(cache)}건 / 주소 매핑: {len(addr_map)}건 / 현대 관리대수 매핑: {len(mgmt_map)}건 / 고유 주소 종류: {len(unit_counts)}개")

    records = []
    skipped_no_addr = 0
    skipped_no_coord = 0

    with pyxlsb.open_workbook(DB_SOURCE) as wb:
        with wb.get_sheet(DB_SHEET) as sheet:
            rows = sheet.rows()
            header = [c.v for c in next(rows)]
            idx = {name: header.index(name) for name in NEEDED_COLUMNS}
            for row in rows:
                if limit is not None and len(records) >= limit:
                    break
                vals = [c.v for c in row]
                eno = normalize_key(vals[idx["ELEVATORNO"]])
                addr = addr_map.get(eno)
                if not addr:
                    skipped_no_addr += 1
                    continue
                coord = cache.get(addr)
                if not coord:
                    skipped_no_coord += 1
                    continue
                name = vals[idx["BULDNM"]]
                mgmt = mgmt_map.get(eno)
                records.append(
                    {
                        "id": pad_elevator_no(eno),
                        "lat": coord["lat"],
                        "lng": coord["lng"],
                        "name": name,
                        "address": addr,
                        "manufacturer": vals[idx["MANUFACTURERNAME"]],
                        "mntCompany": vals[idx["MNTCPNYNM"]],
                        "unitCount": unit_counts.get(addr),
                        "model": vals[idx["ELVTRMODEL"]],
                        "status": vals[idx["ELVTRSTTS"]],
                        "kind": vals[idx["ELVTRKINDNM"]],
                        "firstInstallDate": vals[idx["FRSTINSTALLATIONDE"]],
                        "installDate": vals[idx["INSTALLATIONDE"]],
                        "projectNo": pad_project_no(mgmt["projectNo"]) if mgmt else None,
                        "team": mgmt["team"] if mgmt else None,
                    }
                )

    by_region = defaultdict(list)
    for r in records:
        by_region[region_of(r["address"])].append(r)

    os.makedirs(REGIONS_DIR, exist_ok=True)
    manifest = {}
    for region, items in by_region.items():
        file_name = f"{region}.json"
        with open(os.path.join(REGIONS_DIR, file_name), "w", encoding="utf-8") as f:
            json.dump(items, f, ensure_ascii=False)
        lats = [it["lat"] for it in items]
        lngs = [it["lng"] for it in items]
        # 제조/유지보수 업체를 둘 다 조합해서 세어둔다. 요약 배지(전국/광역 보기)가
        # 지금 켜진 필터 조합에 맞는 개수를 보여주려면 이 교차표가 있어야 한다
        # (제조업체 하나만 껐을 때, 유지보수업체 하나만 껐을 때, 둘 다 껐을 때 등
        # 모든 조합을 지역 데이터 파일을 새로 불러오지 않고도 바로 계산 가능).
        cross_counts = defaultdict(lambda: defaultdict(int))
        mnt_counts = Counter()
        team_counts = Counter()
        for it in items:
            mfg = classify_company(it["manufacturer"])
            mnt = classify_company(it["mntCompany"])
            cross_counts[mfg][mnt] += 1
            mnt_counts[mnt] += 1
            if it["team"]:
                team_counts[it["team"]] += 1
        manifest[region] = {
            "file": f"regions/{file_name}",
            "count": len(items),
            "minLat": min(lats),
            "maxLat": max(lats),
            "minLng": min(lngs),
            "maxLng": max(lngs),
            "oldCount": sum(1 for it in items if is_old(it["firstInstallDate"])),
            "mntCounts": dict(mnt_counts),
            "crossCounts": {mfg: dict(mnt_map) for mfg, mnt_map in cross_counts.items()},
            "teamCounts": dict(team_counts),
        }

    with open(MANIFEST_FILE, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False)

    # 지도 데이터(지역 파일)는 화면에 걸치는 지역만 불러오지만, 검색은 지금 화면과
    # 무관하게 전체에서 찾을 수 있어야 하므로 훨씬 가벼운 검색 전용 색인을 따로 만든다.
    # (좌표/주소/제조업체 등은 빼고 검색에 필요한 필드만 배열로 - 용량을 최소화)
    search_index = [
        [r["id"], r["name"], r["projectNo"], region_of(r["address"])] for r in records
    ]
    with open(SEARCH_INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump(search_index, f, ensure_ascii=False)

    print(f"완료: {len(records)}건 -> {len(by_region)}개 지역 파일 + {MANIFEST_FILE}")
    print(f"검색 색인: {len(search_index)}건 -> {SEARCH_INDEX_FILE}")
    for region, info in sorted(manifest.items(), key=lambda x: -x[1]["count"]):
        print(f"  {region}: {info['count']}건")
    print(f"주소 매핑 없음(스킵): {skipped_no_addr}건")
    print(f"좌표 미완료(아직 지오코딩 안됨, 스킵): {skipped_no_coord}건")
    print(f"프로젝트번호 있는 건(현대엘리베이터 직영): {sum(1 for r in records if r['projectNo'])}건")


if __name__ == "__main__":
    main()
