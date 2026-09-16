"""
건축물대장(표제부) 공통 유틸 - 조회키 만들기와 "여러 동 중 어느 동인가" 매칭 로직.

수집(fetch_building_ledger.py)과 조립(build_dataset.py)이 같은 규칙을 써야 하므로
여기 한 곳에 모아둔다.

핵심 개념 - 왜 매칭이 필요한가:
  건축물대장은 "지번(PNU) + 동" 단위다. 아파트 단지 한 지번에 101동/102동...이
  전부 별개의 표제부로 들어있고, 상가 한 지번에도 여러 동이 있을 수 있다.
  반면 우리 데이터의 현장 식별자는 (주소, 현장명=BULDNM) 뿐이다.
  그래서 지번으로 받아온 표제부 목록 중 우리 현장명에 해당하는 동을 골라내야 한다.

  자체 검증 수단: 표제부의 rideUseElvtCnt(승용승강기수) + emgenUseElvtCnt(비상용승강기수)와
  우리 unitCount(주소+현장명 기준 대수)를 비교하면 매칭이 맞았는지 대략 확인할 수 있다.

실행: python pipeline/building_ledger.py --test   (매칭 로직 자체 검사)
"""
import re
import sys
import unicodedata

# 매칭 신뢰도 - 낮을수록 확실하다. 리포트/필터링에 쓴다.
MATCH_SINGLE = "single"      # 그 지번에 동이 하나뿐 -> 고민할 것 없음
MATCH_DONG = "dong"          # 현장명의 "101동" 같은 동 번호가 대장 동명칭과 일치
MATCH_EXACT = "exact"        # 현장명과 대장의 건물명/동명칭이 (정규화 후) 완전 일치
MATCH_PARTIAL = "partial"    # 한쪽이 다른 쪽에 포함됨
MATCH_LARGEST = "largest"    # 못 고름 -> 주건축물 중 연면적이 가장 큰 동으로 대표
MATCH_NONE = "none"          # 대장 자체가 없음

# 위 순서가 그대로 신뢰도 순서 (앞쪽이 더 믿을 만함)
MATCH_QUALITY_ORDER = [MATCH_SINGLE, MATCH_EXACT, MATCH_DONG, MATCH_PARTIAL, MATCH_LARGEST, MATCH_NONE]

# build_dataset.py의 조인 리포트 출력용
MATCH_QUALITY_LABELS = {
    MATCH_SINGLE: "동 하나뿐(확실)",
    MATCH_EXACT: "건물명/동명칭 완전일치",
    MATCH_DONG: "동 번호 일치",
    MATCH_PARTIAL: "이름 부분일치",
    MATCH_LARGEST: "동 못 고름 -> 대표 건물로 대체(검증 필요)",
    MATCH_NONE: "대장 없음(미수집 또는 자료없음)",
}

# 표제부에서 우리가 실제로 쓰는 필드만 남긴다(응답 필드가 50개가 넘어서 전부 저장하면
# 캐시가 불필요하게 커짐). 값의 의미는 docs/building-ledger.md 참고.
LEDGER_FIELDS = [
    "mgmBldrgstPk",     # 관리건축물대장PK - 층별개요/전유부 등 후속 조회의 키
    "bldNm",            # 건물명
    "dongNm",           # 동명칭
    "mainPurpsCdNm",    # 주용도코드명  <- 우리가 제일 원하는 값
    "etcPurps",         # 기타용도(주용도 밑의 실제 용도 상세)
    "mainAtchGbCdNm",   # 주건축물/부속건축물 구분
    "totArea",          # 연면적(㎡)
    "archArea",         # 건축면적(㎡)
    "grndFlrCnt",       # 지상층수
    "ugrndFlrCnt",      # 지하층수
    "hhldCnt",          # 세대수
    "hoCnt",            # 호수
    "fmlyCnt",          # 가구수
    "useAprDay",        # 사용승인일
    "rideUseElvtCnt",   # 승용승강기수  <- 우리 대수와 교차검증용
    "emgenUseElvtCnt",  # 비상용승강기수
    "strctCdNm",        # 구조
    "platPlc",          # 대지위치(지번주소) - 조인 검증용
    "newPlatPlc",       # 도로명대지위치 - 조인 검증용
]


def pnu_key(jibun):
    """건축물대장 조회 단위 키. 주소 여러 개가 같은 지번을 가리키면 API 호출도 한 번만."""
    return "{}-{}-{}-{}-{}".format(
        jibun["sigunguCd"], jibun["bjdongCd"], jibun["platGbCd"], jibun["bun"], jibun["ji"]
    )


def normalize_name(name):
    """현장명/건물명 비교용 정규화 - 공백, 괄호, 하이픈, 아파트/빌딩 같은 흔한 접미사를 제거."""
    if not name:
        return ""
    text = unicodedata.normalize("NFKC", str(name))
    text = re.sub(r"\(.*?\)", "", text)
    text = re.sub(r"[\s\-_·,.'\"/]", "", text)
    return text.upper()


DONG_PATTERN = re.compile(r"(?:제)?(\d+)\s*동")


def extract_dong_numbers(name):
    """'래미안 제101동', '101동' 등에서 동 번호만 뽑는다."""
    if not name:
        return set()
    return {str(int(m)) for m in DONG_PATTERN.findall(str(name))}


def _dong_numbers_of_ledger(row):
    return extract_dong_numbers(row.get("dongNm")) | extract_dong_numbers(row.get("bldNm"))


def _to_float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _is_main_building(row):
    return "부속" not in str(row.get("mainAtchGbCdNm") or "")


def match_ledger_row(buld_nm, rows):
    """지번으로 받아온 표제부 목록 rows 중 현장명(buld_nm)에 해당하는 한 건을 고른다.
    반환: (선택된 row 또는 None, 매칭 신뢰도 문자열)

    이름으로 맞춰보는 단계(완전일치 -> 동 번호 -> 부분일치)를 부속건축물까지 포함한
    전체 후보에 대해 먼저 돌린다. 현장명이 "관리동"처럼 부속건축물을 직접 가리키는
    경우가 있어서, 주건축물 우선 규칙은 이름으로 못 골랐을 때만 적용한다."""
    if not rows:
        return None, MATCH_NONE
    if len(rows) == 1:
        return rows[0], MATCH_SINGLE

    target = normalize_name(buld_nm)
    target_dongs = extract_dong_numbers(buld_nm)

    # 완전일치는 "그 이름을 가진 행이 딱 하나"일 때만 인정한다. 아파트 단지는
    # 모든 동의 bldNm이 단지명("행복아파트")으로 같아서, 현장명이 단지명뿐이면
    # 어느 동인지 알 수 없는데도 첫 번째 동이 찍혀버리는 문제가 있었다.
    if target:
        exact = [
            row
            for row in rows
            if any(c and normalize_name(c) == target for c in (row.get("dongNm"), row.get("bldNm")))
        ]
        if len(exact) == 1:
            return exact[0], MATCH_EXACT

    # 현장명에 동 번호가 있으면 그게 가장 강한 단서 - "101동"은 101동에만 붙는다.
    if target_dongs:
        dong_hits = [row for row in rows if target_dongs & _dong_numbers_of_ledger(row)]
        if len(dong_hits) == 1:
            return dong_hits[0], MATCH_DONG

    main_rows = [r for r in rows if _is_main_building(r)] or rows
    if len(main_rows) == 1:
        # 건물 한 채 + 주차장동/관리동 같은 부속건축물 조합 - 사실상 고민할 것 없음
        return main_rows[0], MATCH_SINGLE

    # 부분 포함 - "가나아파트" vs "가나아파트제3차" 같은 표기 차이를 흡수한다.
    # 두 가지를 제외해야 엉뚱한 동이 잡히지 않는다:
    #  - 동 번호가 붙은 대장 행: 번호 없는 현장명과 겹치면 아무 동이나 잡힌다.
    #  - 부속건축물의 건물명(bldNm): 단지 전체가 같은 건물명을 쓰므로 현장명이 단지명뿐일 때
    #    관리동/주차장동이 대표로 잡혀버린다. 부속건축물은 동명칭(dongNm)이 현장명에
    #    직접 들어있을 때만 후보로 둔다("행복아파트 관리동" 처럼).
    if target:
        partial = []
        for row in rows:
            if _dong_numbers_of_ledger(row):
                continue
            candidates = [row.get("dongNm")]
            if _is_main_building(row):
                candidates.append(row.get("bldNm"))
            if any(
                c and (normalize_name(c) in target or target in normalize_name(c))
                for c in candidates
            ):
                partial.append(row)
        if len(partial) == 1:
            return partial[0], MATCH_PARTIAL

    # 끝까지 못 고르면 그 지번의 대표 건물(연면적 최대 주건축물)로 둔다.
    # 용도 정보는 같은 지번 안에서 대체로 비슷하므로 아주 틀리지는 않지만,
    # 신뢰도 표시를 남겨서 나중에 걸러낼 수 있게 한다.
    return max(main_rows, key=lambda r: _to_float(r.get("totArea"))), MATCH_LARGEST


def purpose_of(row):
    """화면에 보여줄 용도 문자열. 주용도만으로는 '제2종근린생활시설'처럼 뭉뚱그려지므로
    기타용도(실제 용도 상세)를 같이 붙인다."""
    if not row:
        return None
    main = (row.get("mainPurpsCdNm") or "").strip()
    etc = (row.get("etcPurps") or "").strip()
    if main and etc and etc != main:
        return f"{main}({etc})"
    return main or etc or None


def elevator_count_of(row):
    """대장상 승강기 수(승용+비상용) - 우리 대수와 교차검증용. 값이 없으면 None."""
    if not row:
        return None
    ride = row.get("rideUseElvtCnt")
    emgen = row.get("emgenUseElvtCnt")
    if ride is None and emgen is None:
        return None
    return int(_to_float(ride)) + int(_to_float(emgen))


def _test():
    """API 키 없이 돌려볼 수 있는 매칭 로직 자체 검사."""
    apartment = [
        {"dongNm": "101동", "bldNm": "행복아파트", "mainPurpsCdNm": "공동주택", "etcPurps": "아파트", "totArea": "9000"},
        {"dongNm": "102동", "bldNm": "행복아파트", "mainPurpsCdNm": "공동주택", "etcPurps": "아파트", "totArea": "12000"},
        {"dongNm": "관리동", "bldNm": "행복아파트", "mainPurpsCdNm": "공동주택", "etcPurps": "부대시설",
         "mainAtchGbCdNm": "부속건축물", "totArea": "300"},
    ]
    single = [{"dongNm": "", "bldNm": "예감빌딩", "mainPurpsCdNm": "업무시설", "etcPurps": "사무소", "totArea": "3000"}]
    mixed = [
        {"dongNm": "", "bldNm": "가나빌딩제1차", "mainPurpsCdNm": "제2종근린생활시설", "etcPurps": "소매점", "totArea": "800"},
        {"dongNm": "", "bldNm": "다라상가", "mainPurpsCdNm": "판매시설", "etcPurps": "상점", "totArea": "2500"},
    ]

    cases = [
        ("동 번호 매칭", "행복아파트 101동", apartment, "101동", MATCH_DONG),
        ("동 번호 매칭(제N동 표기)", "행복아파트 제102동", apartment, "102동", MATCH_DONG),
        ("동 하나뿐", "예감빌딩", single, "", MATCH_SINGLE),
        ("현장명이 대장과 달라도 단독이면 채택", "예감빌딩 신관", single, "", MATCH_SINGLE),
        ("건물명 완전일치", "다라상가", mixed, "", MATCH_EXACT),
        ("건물명 부분일치", "가나빌딩", mixed, "", MATCH_PARTIAL),
        ("못 고르면 연면적 최대 주건축물", "이름없는현장", apartment, "102동", MATCH_LARGEST),
        ("현장명이 부속건축물을 직접 가리키면 그 동", "행복아파트 관리동", apartment, "관리동", MATCH_PARTIAL),
        ("건물 1채 + 부속건축물이면 주건축물", "무명현장", mixed[:1] + [dict(mixed[1], mainAtchGbCdNm="부속건축물")], "", MATCH_SINGLE),
        # 단지명만 있으면 어느 동인지 알 수 없다 -> 완전일치로 착각하지 말고 대표 건물로
        ("단지명만 있는 아파트는 동을 특정 못함", "행복아파트", apartment, "102동", MATCH_LARGEST),
    ]

    failures = 0
    for label, buld_nm, rows, expect_dong, expect_quality in cases:
        row, quality = match_ledger_row(buld_nm, rows)
        got_dong = (row or {}).get("dongNm", "")
        ok = got_dong == expect_dong and quality == expect_quality
        print(f"  [{'OK ' if ok else 'FAIL'}] {label}: dongNm={got_dong!r} quality={quality}")
        if not ok:
            print(f"         기대값: dongNm={expect_dong!r} quality={expect_quality}")
            failures += 1

    row, _ = match_ledger_row("행복아파트 101동", apartment)
    assert purpose_of(row) == "공동주택(아파트)", purpose_of(row)
    assert match_ledger_row("아무거나", []) == (None, MATCH_NONE)
    assert elevator_count_of({"rideUseElvtCnt": "2", "emgenUseElvtCnt": "1"}) == 3
    assert elevator_count_of({}) is None
    assert pnu_key({"sigunguCd": "11680", "bjdongCd": "10100", "platGbCd": "0", "bun": "0736", "ji": "0017"}) == "11680-10100-0-0736-0017"
    print("  [OK ] purpose_of / elevator_count_of / pnu_key")

    if failures:
        raise SystemExit(f"\n매칭 로직 검사 실패 {failures}건")
    print("\n매칭 로직 검사 통과")


if __name__ == "__main__":
    if "--test" in sys.argv:
        _test()
    else:
        print(__doc__)
