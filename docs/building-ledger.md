# 건축물대장으로 건물용도 보강하기

현대엘리베이터 승강기가 설치된 현장의 **건물용도**를 제대로 파악하기 위해,
국토교통부 건축물대장 정보를 우리 데이터에 붙이는 방법을 정리한 문서.

## 1. 왜 필요한가 (지금 데이터의 한계)

국가DB에 이미 `BULDPRPOS`(건물용도) 컬럼이 있고, 상세정보 표의 "건물정보"로 보여주고 있다.
그런데 이 값은 **승강기 설치 신고 시점에 적은 값**이라 실제 영업에 쓰기엔 너무 뭉뚱그려져 있다.

전국 886,067건 기준 실제 분포 (고유값 61종):

| BULDPRPOS | 건수 |
| --- | --- |
| 공동주택-분양 | 396,056 |
| 제2종근린생활시설-**기타** | 87,034 |
| 단독주택-다가구주택 | 49,351 |
| 업무시설-**기타** | 49,203 |
| 제1종근린생활시설-**기타** | 46,534 |
| 공장-공장 | 39,663 |
| 교육연구시설-**기타** | 26,247 |

"-기타"로 끝나는 것만 25만 건이 넘는다. 이게 병원인지 사무실인지 학원인지 모텔인지 구분이 안 된다.
건축물대장의 **주용도(`mainPurpsCdNm`) + 기타용도(`etcPurps`)** 를 가져오면
`제2종근린생활시설(학원)`, `업무시설(오피스텔)` 처럼 실제 용도까지 내려간다.

덤으로 같이 따라오는 값들 — 영업 타깃 선별에 바로 쓸 수 있다:

- `useAprDay` 사용승인일 — 건물 노후도 (승강기 교체/리모델링 영업)
- `totArea` 연면적, `grndFlrCnt`/`ugrndFlrCnt` 지상/지하 층수 — 현장 규모
- `hhldCnt` 세대수 — 공동주택 규모
- `rideUseElvtCnt` + `emgenUseElvtCnt` **대장상 승강기 수** — 우리 관리대수와 비교하면
  "대장에는 6대인데 우리가 아는 건 4대" 같은 **미확보 물량**이 바로 드러난다
- `mgmBldrgstPk` 관리건축물대장PK — 층별개요/전유부 등 후속 조회의 키

## 2. 조인 키 — 이게 핵심 질문의 답

### 결론부터

**승강기번호 ↔ 건축물대장을 직접 잇는 공통키는 없다.**
건축물대장은 승강기번호를 모르고, 승강기 국가DB는 건물관리번호를 갖고 있지 않다.
그래서 **주소를 중간 다리로 삼아 2단 조인**한다.

```
[지역별 관리대수(원본)]                    ← 현대엘리베이터 자체 데이터
        │  승강기번호 (= 원PJT와 같은 행)
        │  ★ 이미 구현됨: load_hyundai_management_map()
        ▼
[국가DB aa 시트]  ELEVATORNO ─┬─ BULDNM (현장명)
        │                     └─ BULDPRPOS 등
        │  ELEVATORNO
        ▼
[국가DB_SID_주소.xlsx]  ADDRESS1 (도로명주소)
        │
        │  ▼ 도로명주소 → 지번 변환 (카카오 주소검색 API)  ★ resolve_jibun.py
        ▼
[지번 키]  sigunguCd(5) + bjdongCd(5) + platGbCd(1) + bun(4) + ji(4)
        │            = 사실상 PNU 19자리
        │  ▼ 건축물대장 표제부 조회                        ★ fetch_building_ledger.py
        ▼
[건축물대장 표제부]  동이 여러 개면 여러 건
        │
        │  ▼ BULDNM ↔ dongNm/bldNm 매칭                   ★ match_ledger_row()
        ▼
   주용도 / 사용승인일 / 연면적 / 층수 / 대장 승강기 수
```

### 단계별로

| 단계 | 조인 키 | 상태 |
| --- | --- | --- |
| 관리대수 ↔ 국가DB | `승강기번호` = `ELEVATORNO` | **이미 동작 중** (프로젝트번호/팀 채우는 그 경로) |
| 국가DB ↔ 주소 | `ELEVATORNO` | **이미 동작 중** (`load_addr_map()`) |
| 주소 ↔ 건축물대장 | `sigunguCd`+`bjdongCd`+`platGbCd`+`bun`+`ji` | 신규 — `resolve_jibun.py` |
| 지번 ↔ 동(棟) | `BULDNM` ↔ `dongNm`/`bldNm` | 신규 — `match_ledger_row()` |

즉 **"지역별 관리대수의 값과 건축물대장을 조인하는 키"는 `승강기번호 → 주소 → 지번(PNU)`** 이다.
관리대수 파일에 있는 승강기번호가 출발점이고, 건축물대장 쪽 도착점은 지번이다.

### 왜 주소를 그대로 못 쓰는가

건축물대장 API는 **도로명주소로 조회할 수 없다.** 지번 기반 파라미터만 받는다.
그런데 우리 `ADDRESS1`은 전부 도로명주소다 (전수 확인함).

```
'서울특별시 강남구 선릉로90길 40'   | '예감빌딩'
'경기도 수원시 팔달구 팔달로165번길 51' | '대진빌리지A동'
```

그래서 도로명주소 → 지번 변환이 반드시 한 번 들어가야 한다.

### 변환은 추가 API 신청 없이 된다

이미 지오코딩에 쓰고 있는 **카카오 주소검색 API 응답 안에 지번이 같이 들어있다.**

```jsonc
// GET https://dapi.kakao.com/v2/local/search/address.json?query=서울특별시 강남구 선릉로90길 40
"address": {
  "address_name": "서울 강남구 역삼동 736-17",
  "b_code": "1168010100",      // 법정동코드 10자리
  "main_address_no": "736",    // 본번
  "sub_address_no": "17",      // 부번
  "mountain_yn": "N"           // 산 여부
}
```

매핑 규칙:

| 건축물대장 파라미터 | 카카오 응답에서 만드는 법 | 예 |
| --- | --- | --- |
| `sigunguCd` | `b_code[0:5]` | `11680` (강남구) |
| `bjdongCd` | `b_code[5:10]` | `10100` (역삼동) |
| `platGbCd` | `mountain_yn == "Y"` → `1`, 아니면 `0` (0=대지, 1=산, 2=블록) | `0` |
| `bun` | `main_address_no` 4자리 0패딩 | `0736` |
| `ji` | `sub_address_no` 4자리 0패딩 (없으면 `0000`) | `0017` |

도로명주소가 너무 새로워서 `address`가 비어 오는 경우가 있어,
그때는 **이미 캐시에 있는 위경도**로 좌표→주소 변환(`geo/coord2address`)을 한 번 더 시도한다.

### 남은 난제: 한 지번에 동이 여러 개

건축물대장은 "지번 + 동" 단위다. 아파트 단지 한 지번에 101동/102동…이 전부 별개 표제부로 들어있다.
우리 쪽 현장 식별자는 `(주소, 현장명)` 뿐이므로 어느 동인지 골라내야 한다.
`building_ledger.match_ledger_row()`가 이 순서로 판단하고, **판단 근거를 신뢰도로 같이 남긴다**:

| 신뢰도 | 판단 | 믿을 만한가 |
| --- | --- | --- |
| `single` | 그 지번에 동이 하나뿐 (또는 주건축물이 하나뿐) | 확실 |
| `exact` | 현장명 = 동명칭/건물명, 그런 행이 딱 하나 | 확실 |
| `dong` | 현장명의 "101동"이 대장 동명칭과 일치 | 확실 |
| `partial` | 한쪽 이름이 다른 쪽에 포함 (`가나빌딩` ↔ `가나빌딩제1차`) | 대체로 맞음 |
| `largest` | **못 고름** → 그 지번의 연면적 최대 주건축물로 대체 | 검증 필요 |
| `none` | 대장 자체가 없음 (미수집 또는 자료없음) | — |

`largest` 비율이 높게 나오면 매칭 규칙을 보강해야 한다는 신호다.
`build_dataset.py`가 실행 끝에 이 분포를 출력하고, 지도 화면에서도 `largest`인 경우
"(동 미확정 - 지번 대표 건물 기준)"이라고 표시한다.

자체 검증 수단: 대장의 `rideUseElvtCnt + emgenUseElvtCnt` 와 우리 `unitCount`를 비교하면
매칭이 맞았는지 대략 확인된다. 단, 두 값이 다른 게 꼭 오류는 아니다 (위 "미확보 물량" 참고).

주의: 아파트 단지 현장명이 단지명뿐이면(`행복아파트`) 동을 특정할 수 없다.
모든 동의 `bldNm`이 단지명으로 같기 때문. 이때는 `exact`로 착각하지 않고 `largest`로 떨어진다.

## 3. API 정보

- 데이터셋: **국토교통부_건축HUB_건축물대장정보 서비스** (공공데이터포털 `15134735`)
- 엔드포인트: `https://apis.data.go.kr/1613000/BldRgstHubService/getBrTitleInfo` (표제부)
  - 구버전 `BldRgstService_v2` 도 아직 살아있지만 신규는 `BldRgstHubService` 를 쓴다
- 다른 오퍼레이션: `getBrBasisOulnInfo`(기본개요), `getBrRecapTitleInfo`(총괄표제부),
  `getBrFlrOulnInfo`(층별개요), `getBrExposPubuseAreaInfo`(전유공용면적), `getBrJijiguInfo`(지역지구구역) 등
- 요청 파라미터: `serviceKey`, `sigunguCd`, `bjdongCd`, `platGbCd`, `bun`, `ji`,
  `numOfRows`, `pageNo`, `_type=json`
- 인증키는 공공데이터포털 마이페이지의 **일반 인증키(Decoding)** 값을 쓴다
  (Encoding 값을 넣어도 `fetch_building_ledger.py`가 알아서 되돌린다)

표제부 대신 **총괄표제부**(`getBrRecapTitleInfo`)를 쓰면 단지 전체가 1건으로 와서 동 매칭이 필요 없지만,
동별 용도/층수/승강기 수가 사라진다. 우리는 동별 정보가 필요하므로 표제부를 기본으로 한다.

## 4. 호출량 계산

조회는 **지번 단위**라, 같은 지번의 주소/승강기가 여러 개면 호출은 한 번이다.

- 국가DB 승강기: 886,067건
- 고유 주소: 약 35만 건 (지오코딩 캐시 기준 — 실제 수치는 `resolve_jibun.py` 실행 시 출력됨)
- 고유 지번: 주소보다 조금 적음 (같은 대지에 도로명주소가 여러 개인 경우)

| 계정 | 일 한도 | 전국 소요 |
| --- | --- | --- |
| 개발계정 | 1,000건 | 1년 이상 (비현실적) |
| 운영계정 | 100,000건 (승인 후) | 약 4일 |

→ **운영계정 증량 신청이 사실상 필수.** 신청 전에는 `--region` 으로 파일럿 지역만 먼저 받는다.
이 저장소가 `DETAIL_PILOT_REGIONS` / `TERRITORY_PILOT_REGIONS` 로 이미 쓰고 있는 방식과 같다.

## 5. 실행 순서

```bash
# 0) 인증키 동작 확인 (1건만 호출)
python pipeline/fetch_building_ledger.py --self-test

# 1) 도로명주소 -> 지번 변환 (카카오, 하루 9만 건씩)
python pipeline/resolve_jibun.py

# 2) 건축물대장 표제부 수집 - 먼저 파일럿 지역만
python pipeline/fetch_building_ledger.py --region "서울특별시 강남구" --limit 900

# 3) 지도 데이터 생성 (대장 캐시가 있으면 자동으로 붙는다)
python pipeline/build_dataset.py

# 4) 전국 확장 (운영계정 증량 승인 후)
python pipeline/fetch_building_ledger.py
```

캐시 파일은 모두 `pipeline/cache/` 아래에 쌓이고 `.gitignore` 대상이다.
중간에 끊겨도 다음 실행 때 이어서 진행한다.

| 파일 | 내용 |
| --- | --- |
| `jibun_cache.json` | 주소 → 지번 조회키 |
| `jibun_failed.json` | 지번 변환 실패 주소 |
| `building_ledger.json` | 지번키 → 표제부 목록 (자료없음도 `[]`로 기록해 재조회 방지) |
| `building_ledger_failed.json` | 조회 실패 지번 + 오류 메시지 |

## 6. 데이터 저장 위치

`data/details/{지역}.json` 의 각 행에 7번 인덱스부터 대장 값이 붙는다 (기존 0~6번은 그대로).

```
[승강기번호, 건물정보, 구분, TM, 인승, 속도, 정지층수,
 대장용도, 사용승인일, 연면적, 지상층수, 지하층수, 대장승강기수, 매칭신뢰도]
```

지도 마커용 `data/regions/{지역}.json` 에는 넣지 않는다. 88만 건 전체에 필드가 반복되면
지도 탐색이 느려지므로, 상세정보 버튼을 누른 사람만 해당 지역 파일을 받는 기존 방식을 그대로 따른다.
대장 값은 **건물 단위**이므로 화면에서는 표 안에 대수만큼 반복하지 않고 표 위에 한 줄 요약으로 보여준다.

> 파일 크기 주의: 같은 건물의 승강기 행마다 대장 값이 중복 저장된다.
> 파일럿 지역에서 실제 증가량을 재본 뒤 전국 확장을 판단할 것. 너무 커지면
> 용도 문자열을 `search_index.json` 처럼 사전 인코딩(고유값 테이블 + 인덱스)으로 바꾸면 된다.

## 7. 아직 확인이 필요한 것

1. **국가DB 원본에 지번주소 컬럼이 있는지** (`ADDRESS2` 등). 있으면 카카오 변환 단계를
   건너뛸 수 있어서 35만 건 API 호출이 사라진다. → `국가DB_260805.xlsb` 의 `aa` 시트 헤더 확인
2. **공공데이터포털 운영계정 증량 승인 여부** — 개발계정(1,000건/일)이면 파일럿만 가능
3. **필요한 필드 범위** — 지금은 용도/사용승인일/연면적/층수/승강기수. 세대수·구조·주차대수 등이
   영업에 더 필요하면 `building_ledger.LEDGER_FIELDS` 에 추가하면 된다
4. **파일럿 지역** — 기존 파일럿과 같이 강남구/서초구로 갈지, 다른 지역으로 할지
5. **건물용도로 지도 필터링이 필요한지** — 필요하면 용도를 `data/regions/` 쪽에도 넣고
   `regions_manifest.json` 에 용도별 집계를 추가해야 한다 (현재는 상세정보 표시 전용)

## 참고

- [국토교통부_건축HUB_건축물대장정보 서비스 (공공데이터포털)](https://www.data.go.kr/data/15134735/openapi.do)
- [건축물대장 표제부 (건축HUB)](https://www.hub.go.kr/portal/opn/tyb/idx-bdrg-ttlldr.do)
- [PublicDataReader - 건축물대장정보 사용 가이드](https://github.com/WooilJeong/PublicDataReader/blob/main/assets/docs/portal/BuildingLedger.md)
- [카카오 로컬 API - 주소 검색](https://developers.kakao.com/docs/latest/ko/local/dev-guide#search-by-address)
