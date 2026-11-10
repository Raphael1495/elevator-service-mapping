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
# 기존(직전 DB 시점) 승강기의 주소 - 행정구역 개편(전남광주통합특별시, 인천 검단구/서해구 등)
# 으로 새 DB의 주소 표기가 바뀌었어도 기존 승강기는 이 주소와 좌표를 그대로 쓴다.
ADDRESS_SOURCE = os.path.join(BASE_DIR, "국가DB_SID_주소_260805.xlsx")
# 메인 데이터는 국가DB_261105.xlsb. 시트명이 바뀌어도(aa -> 승강기정보센터DB) 동작하도록 첫 시트를 쓴다.
DB_SOURCE = os.path.join(BASE_DIR, "국가DB_261105.xlsb")
DB_SHEET = 1
MGMT_SOURCE = os.path.join(BASE_DIR, "2026.10월 관리대수(원본).xlsb")
MGMT_SHEET = "리스트"
ORG_FILE = os.path.join(BASE_DIR, "data", "org_hierarchy.json")
REGIONS_DIR = os.path.join(BASE_DIR, "data", "regions")
MANIFEST_FILE = os.path.join(BASE_DIR, "data", "regions_manifest.json")
SEARCH_INDEX_FILE = os.path.join(BASE_DIR, "data", "search_index.json")
DETAILS_DIR = os.path.join(BASE_DIR, "data", "details")
TERRITORY_FILE = os.path.join(BASE_DIR, "data", "territory_companies.json")

# "업체 영역 보기" 기능 - 구(지역) 안에서 유의미하게 활동 중인 업체 목록.
# 대형 4사는 지사 상관없이 통합해서 하나로, 나머지(기타)는 원본 업체명 그대로
# 개별 집계 후 상위 N개만(노이즈 방지로 최소 대수 조건도 같이 검) 남긴다.
# 강남구/서초구 파일럿 검증 후 전국 확장 예정.
TERRITORY_PILOT_REGIONS = {"서울특별시 강남구", "서울특별시 서초구"}
TERRITORY_TOP_ETC = 7
TERRITORY_MIN_COUNT = 10
TERRITORY_PLACEHOLDER_NAMES = {"보수안함", "자체보수"}

# 상세정보(정보창의 "상세정보" 버튼) 전용 필드. 지도 마커용 지역 데이터에 그대로
# 넣으면 88만여 건 전체에 필드 5개가 반복돼 용량이 크게 늘어나므로, 별도의
# 지연로딩 파일(data/details/{지역}.json)로 분리한다. 제주도 파일럿 검증 후
# 전국 확장(예상 총 용량 약 90MB, 지역별로 잘게 나뉘어 있어 개별 파일 크기는
# 문제 없음).
DETAIL_PILOT_REGIONS = None

NEEDED_COLUMNS = [
    "ELEVATORNO",
    "ADDRESS1",
    "BULDNM",
    "MANUFACTURERNAME",
    "MNTCPNYNM",
    "ELVTRSTTS",
    "ELVTRKINDNM",
    "ELVTRMODEL",
    "INSTALLATIONDE",
    "FRSTINSTALLATIONDE",
    "BULDPRPOS",
    "ELVTRDIVNM",
    "ELVTRFORMNM",
    "RATEDCAP",
    "RATEDSPEED",
    "SHUTTLEFLOORCNT",
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


# 신규 승강기의 새 행정구역 표기 -> 기존(개편 전) 표기. 기존 승강기는 기존 주소를 그대로
# 쓰므로, 한 건물의 기존/신규 승강기가 서로 다른 지역·현장으로 갈라지지 않게 맞춘다.
GWANGJU_GU = {"동구", "서구", "남구", "북구", "광산구"}
INCHEON_GU_TO_LEGACY = {"검단구": "서구", "서해구": "서구", "영종구": "중구", "제물포구": "중구"}


def to_legacy_address(address):
    tokens = address.split(" ")
    if len(tokens) > 1:
        if tokens[0] == "전남광주통합특별시":
            tokens[0] = "광주광역시" if tokens[1] in GWANGJU_GU else "전라남도"
        elif SIDO_ALIASES.get(tokens[0], tokens[0]) == "인천광역시" and tokens[1] in INCHEON_GU_TO_LEGACY:
            tokens[1] = INCHEON_GU_TO_LEGACY[tokens[1]]
    return " ".join(tokens)


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
    합쳐지는 문제가 있었음 -> 주소(사실상 유일한 건물 식별자) 기준으로 변경.
    (참고용/export_sample.py 전용 - 실제 지도 데이터는 아래 count_units_by_address_and_name 사용)"""
    return Counter(addr_map.values())


def count_units_by_address_and_name(pairs):
    """같은 주소 + 같은 현장명(BULDNM)이어야 하나의 현장으로 묶어서 대수를 센다.
    주소만으로 묶으면 한 주소에 여러 동/시설(현장명이 다른 경우)이 있을 때 실제로는
    별개인 현장이 하나로 합쳐지는 문제가 있어, 현장명까지 같아야 같은 현장으로 취급."""
    return Counter((addr, (name or "").strip()) for addr, name in pairs)


def load_cache():
    if not os.path.exists(CACHE_FILE):
        raise SystemExit(f"{CACHE_FILE} 이 없습니다. 먼저 pipeline/geocode.py를 실행하세요.")
    with open(CACHE_FILE, encoding="utf-8") as f:
        return json.load(f)


def load_addr_map():
    df = pd.read_excel(ADDRESS_SOURCE)
    df["ADDRESS1"] = df["ADDRESS1"].astype(str).str.strip()
    return {normalize_key(k): v for k, v in zip(df["ELEVATORNO"], df["ADDRESS1"])}


MGMT_TYPE_COLUMN_CANDIDATES = ["직영/협력사", "직영협력사", "관리구분", "직영/협력"]


def _find_column(header, candidates):
    """정확히 일치하는 컬럼명을 먼저 찾고, 없으면 후보 키워드가 포함된 컬럼을 찾는다.
    관리대수 원본 파일의 실제 컬럼명이 조금 다를 수 있어 유연하게 대응."""
    for cand in candidates:
        if cand in header:
            return header.index(cand)
    for i, h in enumerate(header):
        if isinstance(h, str) and any(cand in h for cand in candidates):
            return i
    return None


# 담당구분(DK) -> 화면 표시용 짧은 이름. 목록 순서가 곧 화면 표시 순서다.
DAM_MAP = {
    "서비스강북담당": "강북", "서비스강남담당": "강남", "서비스동부담당": "동부",
    "서비스서부담당": "서부", "ES담당": "ES", "자회사": "자회사",
}
DAM_ORDER = ["강북", "강남", "동부", "서부", "ES", "자회사"]
# 팀(지사/센터) 표시 순서 - 목록에 없으면 이름순으로 뒤에 붙는다.
TEAM_ORDER = [
    "강북지사(서비스)", "강서지사(서비스)", "인천지사(서비스)",
    "강남지사(서비스)", "분당지사(서비스)", "수원지사(서비스)", "강원지사(서비스)",
    "대구지사(서비스)", "부산지사(서비스)", "울산지사(서비스)", "경남지사(서비스)",
    "대전지사(서비스)", "충청지사(서비스)", "광주지사(서비스)", "전주지사(서비스)", "제주지사(서비스)",
    "ES서비스팀",
]
TYPE_ORDER = ["직영", "협력사", "자회사"]

# 10월 관리대수 원본(리스트 시트)에서 반드시 있어야 하는 컬럼. 컬럼 위치가 아니라 헤더 이름으로
# 찾고, 사용자가 안내한 엑셀 열 위치(X=팀, AA=직영/협력사, DI=SID, DK=담당구분)와 다르면 경고한다.
MGMT_EXPECTED_LETTERS = {"팀": "X", "파트너사": "Z", "직영/협력사": "AA", "SID": "DI", "담당구분": "DK"}


def _col_letter(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _clean(v):
    return v.strip() if isinstance(v, str) and v.strip() else None


def load_hyundai_management_map():
    """현대엘리베이터 자체 관리대수(원본, 리스트 시트)에서
    승강기번호(SID) -> {프로젝트번호, 담당(dam), 팀(team), 직영구분(mgmtType), 파트너사(partner)} 매핑.
    SID가 공란/숫자 아님/0000000(자리표시값)인 행은 조인할 수 없으니 버린다.
    같은 SID가 서로 다른 소속으로 여러 번 나오면 나중 행을 쓰고 목록을 출력한다."""
    if not os.path.exists(MGMT_SOURCE):
        print(f"경고: {MGMT_SOURCE} 없음 - 프로젝트번호/담당팀 없이 진행")
        return {}
    result = {}
    skipped = Counter()
    conflicts = []
    with pyxlsb.open_workbook(MGMT_SOURCE) as wb:
        with wb.get_sheet(MGMT_SHEET) as sheet:
            rows = sheet.rows()
            header = [c.v for c in next(rows)]
            idx = {name: header.index(name) for name in
                   ("원PJT", "팀", "파트너사", "직영/협력사", "SID", "담당구분")}
            for name, letter in MGMT_EXPECTED_LETTERS.items():
                actual = _col_letter(idx[name])
                if actual != letter:
                    print(f"경고: 관리대수 원본의 '{name}' 컬럼이 {letter}가 아니라 {actual} 열에 있음 - 헤더 이름 기준으로 계속 진행")
            for row in rows:
                vals = [c.v for c in row]
                vals += [None] * (len(header) - len(vals))
                raw_sid = vals[idx["SID"]]
                sid = str(raw_sid).strip() if raw_sid is not None else ""
                if not sid.isdigit():
                    skipped["SID 공란/숫자아님"] += 1
                    continue
                eno = normalize_key(sid)
                if int(eno) == 0:
                    skipped["SID 0000000"] += 1
                    continue
                dam = DAM_MAP.get(_clean(vals[idx["담당구분"]]))
                team = _clean(vals[idx["팀"]])
                mgmt_type = _clean(vals[idx["직영/협력사"]])
                if not (dam and team and mgmt_type):
                    skipped["담당/팀/구분 누락"] += 1
                    continue
                pjt = vals[idx["원PJT"]]
                entry = {
                    "projectNo": normalize_key(pjt) if pjt is not None else None,
                    "dam": dam,
                    "team": team,
                    "mgmtType": mgmt_type,
                    "partner": _clean(vals[idx["파트너사"]]),
                }
                prev = result.get(eno)
                if prev and (prev["dam"], prev["team"], prev["mgmtType"], prev["partner"]) != (
                        entry["dam"], entry["team"], entry["mgmtType"], entry["partner"]):
                    conflicts.append((eno, prev["team"], prev["partner"], entry["team"], entry["partner"]))
                result[eno] = entry
    if skipped:
        print("관리대수 원본에서 제외한 행: " + ", ".join(f"{k} {v}건" for k, v in skipped.items()))
    if conflicts:
        print(f"같은 SID가 다른 소속으로 중복된 건 {len(conflicts)}건(나중 행 사용):")
        for c in conflicts:
            print(f"  {c[0]}: {c[1]}/{c[2]} -> {c[3]}/{c[4]}")
    return result


def build_org_hierarchy(records):
    """records의 _org(담당, 팀, 구분, 파트너사)에서 고유 노드 목록을 만들고 각 레코드에 pt(노드 번호)를
    붙인다. 화면 필터(담당 > 팀 > 구분 > 파트너사)가 이 표를 기준으로 단계별 목록을 만든다."""
    nodes = sorted(
        {r["_org"] for r in records if r.get("_org")},
        key=lambda n: (
            DAM_ORDER.index(n[0]) if n[0] in DAM_ORDER else 99,
            TEAM_ORDER.index(n[1]) if n[1] in TEAM_ORDER else 99, n[1],
            TYPE_ORDER.index(n[2]) if n[2] in TYPE_ORDER else 99,
            n[3] or "",
        ),
    )
    index = {n: i for i, n in enumerate(nodes)}
    for r in records:
        org = r.pop("_org", None)
        r["pt"] = index[org] if org else None
    return [list(n) for n in nodes]


def main():
    limit = int(sys.argv[1]) if len(sys.argv) > 1 else None

    cache = load_cache()
    addr_map = load_addr_map()
    mgmt_map = load_hyundai_management_map()
    print(f"지오코딩 캐시: {len(cache)}건 / 주소 매핑: {len(addr_map)}건 / 현대 관리대수 매핑: {len(mgmt_map)}건")

    # 대수(unitCount)는 주소+현장명이 둘 다 같아야 같은 현장으로 묶어서 세야 하는데,
    # 현장명(BULDNM)이 국가DB(DB_SOURCE)에만 있어서, 전체를 한 번 읽어 (승강기번호, 주소,
    # 현장명, 원본값) 목록으로 모아둔 뒤 그 목록으로 카운트와 최종 레코드를 둘 다 만든다
    # (xlsb를 두 번 읽지 않기 위함).
    raw_rows = []
    skipped_no_addr = 0
    with pyxlsb.open_workbook(DB_SOURCE) as wb:
        with wb.get_sheet(DB_SHEET) as sheet:
            rows = sheet.rows()
            header = [c.v for c in next(rows)]
            idx = {name: header.index(name) for name in NEEDED_COLUMNS}
            for row in rows:
                if limit is not None and len(raw_rows) >= limit:
                    break
                vals = [c.v for c in row]
                eno = normalize_key(vals[idx["ELEVATORNO"]])
                new_addr = str(vals[idx["ADDRESS1"]]).strip() if vals[idx["ADDRESS1"]] is not None else ""
                addr = addr_map.get(eno)
                if not addr:
                    if not new_addr:
                        skipped_no_addr += 1
                        continue
                    # 직전 DB에 없던 신규 승강기: 새 DB 주소를 기존 행정구역 표기로 바꿔서 쓴다.
                    # 좌표는 새 표기 그대로 지오코딩해 둔 캐시에서 찾는다.
                    addr = to_legacy_address(new_addr)
                name = vals[idx["BULDNM"]]
                raw_rows.append((eno, addr, name, vals, new_addr))

    unit_counts = count_units_by_address_and_name((addr, name) for _, addr, name, _, _ in raw_rows)
    print(f"고유 현장(주소+현장명) 종류: {len(unit_counts)}개")

    records = []
    detail_by_region = defaultdict(list)
    skipped_no_coord = 0
    for eno, addr, name, vals, new_addr in raw_rows:
        # 기존 주소의 좌표가 없으면(예전 변환 실패) 새 DB 주소로 이미 받아 둔 좌표가 있는지 본다.
        coord = cache.get(addr) or cache.get(new_addr)
        if not coord:
            skipped_no_coord += 1
            continue
        mgmt = mgmt_map.get(eno)
        padded_id = pad_elevator_no(eno)
        records.append(
            {
                "id": padded_id,
                "lat": coord["lat"],
                "lng": coord["lng"],
                "name": name,
                "address": addr,
                "manufacturer": vals[idx["MANUFACTURERNAME"]],
                "mntCompany": vals[idx["MNTCPNYNM"]],
                "unitCount": unit_counts.get((addr, (name or "").strip())),
                "model": vals[idx["ELVTRMODEL"]],
                "status": vals[idx["ELVTRSTTS"]],
                "kind": vals[idx["ELVTRKINDNM"]],
                "firstInstallDate": vals[idx["FRSTINSTALLATIONDE"]],
                "installDate": vals[idx["INSTALLATIONDE"]],
                "projectNo": pad_project_no(mgmt["projectNo"]) if mgmt else None,
                "team": mgmt["team"] if mgmt else None,
                "mgmtType": mgmt["mgmtType"] if mgmt else None,
                "_org": (mgmt["dam"], mgmt["team"], mgmt["mgmtType"], mgmt["partner"]) if mgmt else None,
            }
        )

        region = region_of(addr)
        if DETAIL_PILOT_REGIONS is None or region in DETAIL_PILOT_REGIONS:
            detail_by_region[region].append(
                [
                    padded_id,
                    vals[idx["BULDPRPOS"]],
                    vals[idx["ELVTRDIVNM"]],
                    vals[idx["ELVTRFORMNM"]],
                    vals[idx["RATEDCAP"]],
                    vals[idx["RATEDSPEED"]],
                    vals[idx["SHUTTLEFLOORCNT"]],
                ]
            )

    org_nodes = build_org_hierarchy(records)
    with open(ORG_FILE, "w", encoding="utf-8") as f:
        json.dump({"nodes": org_nodes}, f, ensure_ascii=False)
    print(f"담당/팀/구분/파트너사 계층: 노드 {len(org_nodes)}개 -> {ORG_FILE}")

    by_region = defaultdict(list)
    for r in records:
        by_region[region_of(r["address"])].append(r)

    os.makedirs(REGIONS_DIR, exist_ok=True)
    manifest = {}
    territory_data = {}
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
        mgmt_type_counts = Counter()
        pt_counts = Counter()
        for it in items:
            mfg = classify_company(it["manufacturer"])
            mnt = classify_company(it["mntCompany"])
            cross_counts[mfg][mnt] += 1
            mnt_counts[mnt] += 1
            if it["team"]:
                team_counts[it["team"]] += 1
            if it["mgmtType"]:
                mgmt_type_counts[it["mgmtType"]] += 1
            if it["pt"] is not None:
                pt_counts[str(it["pt"])] += 1
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
            "mgmtTypeCounts": dict(mgmt_type_counts),
            "teamCounts": dict(team_counts),
            "ptCounts": dict(pt_counts),
        }

        if region in TERRITORY_PILOT_REGIONS:
            raw_counts = Counter()
            for it in items:
                mnt_name = (it["mntCompany"] or "").strip()
                if mnt_name and mnt_name not in TERRITORY_PLACEHOLDER_NAMES:
                    raw_counts[mnt_name] += 1
            big4_entries = [
                {"name": g, "count": mnt_counts[g], "type": "big4"}
                for g in ("현대", "오티스", "티케이", "미쓰비시")
                if mnt_counts.get(g)
            ]
            etc_entries = sorted(
                (
                    {"name": name, "count": c, "type": "etc"}
                    for name, c in raw_counts.items()
                    if classify_company(name) == "기타" and c >= TERRITORY_MIN_COUNT
                ),
                key=lambda e: -e["count"],
            )[:TERRITORY_TOP_ETC]
            combined = sorted(big4_entries + etc_entries, key=lambda e: -e["count"])
            territory_data[region] = combined

    with open(MANIFEST_FILE, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False)

    with open(TERRITORY_FILE, "w", encoding="utf-8") as f:
        json.dump(territory_data, f, ensure_ascii=False)
    if territory_data:
        print(f"업체 영역 파일럿: {', '.join(sorted(territory_data.keys()))}")

    # 정보창 "상세정보" 버튼 전용 데이터 - [승강기번호, 건물정보, 구분, TM, 인승, 속도,
    # 정지층수] 압축 배열. 지도 마커 데이터와 분리된 파일이라 버튼을 누른 사람만
    # 필요할 때 딱 그 지역 파일만 받는다(평소 지도 탐색 속도에는 영향 없음).
    os.makedirs(DETAILS_DIR, exist_ok=True)
    for region, rows in detail_by_region.items():
        with open(os.path.join(DETAILS_DIR, f"{region}.json"), "w", encoding="utf-8") as f:
            json.dump(rows, f, ensure_ascii=False)
    if detail_by_region:
        pilot_desc = "전체" if DETAIL_PILOT_REGIONS is None else ", ".join(sorted(detail_by_region.keys()))
        print(f"상세정보 파일: {len(detail_by_region)}개 지역 ({pilot_desc})")

    # 지도 데이터(지역 파일)는 화면에 걸치는 지역만 불러오지만, 검색은 지금 화면과
    # 무관하게 전체에서 찾을 수 있어야 하므로 훨씬 가벼운 검색 전용 색인을 따로 만든다.
    # (좌표/주소/제조업체 등은 빼고 검색에 필요한 필드만 배열로 - 용량을 최소화)
    #
    # 지역명(379종)과 현장명(26만여 종)은 레코드 수(87만+)에 비해 종류가 훨씬 적어
    # 매번 문자열을 반복 저장하면 낭비가 크다(용량의 절반 이상). 고유값을 별도
    # 테이블로 빼고 각 레코드에는 인덱스 번호만 저장 -> 61MB대에서 33MB대로 감소.
    region_names = sorted(by_region.keys())
    region_index = {r: i for i, r in enumerate(region_names)}
    name_values = sorted({r["name"] for r in records}, key=lambda x: (x is None, x))
    name_index = {n: i for i, n in enumerate(name_values)}

    search_items = [
        [r["id"], name_index[r["name"]], r["projectNo"], region_index[region_of(r["address"])]]
        for r in records
    ]
    with open(SEARCH_INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump({"regions": region_names, "names": name_values, "items": search_items}, f, ensure_ascii=False)

    print(f"완료: {len(records)}건 -> {len(by_region)}개 지역 파일 + {MANIFEST_FILE}")
    print(f"검색 색인: {len(search_items)}건 (지역 {len(region_names)}종, 현장명 {len(name_values)}종) -> {SEARCH_INDEX_FILE}")
    for region, info in sorted(manifest.items(), key=lambda x: -x[1]["count"]):
        print(f"  {region}: {info['count']}건")
    print(f"주소 매핑 없음(스킵): {skipped_no_addr}건")
    print(f"좌표 미완료(아직 지오코딩 안됨, 스킵): {skipped_no_coord}건")
    print(f"프로젝트번호 있는 건(현대엘리베이터 직영): {sum(1 for r in records if r['projectNo'])}건")


if __name__ == "__main__":
    main()
