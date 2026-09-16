# 승강기 서비스 영업현장 맵핑

전국 승강기 설치 위치를 카카오맵에 찍어 서비스 보수 영업에 활용하기 위한 프로젝트.

## 폴더 구조

```
index.html                     지도 페이지 (카카오맵 + 마커 클러스터링)
data/elevators.json            지도가 읽는 실데이터 (pipeline/build_dataset.py가 생성, git 커밋 대상)
data/elevators.sample.json     실데이터 없을 때 보여줄 샘플 5건
pipeline/geocode.py            주소 -> 위경도 지오코딩 (카카오 로컬 API, 하루 한도만큼씩 실행)
pipeline/resolve_jibun.py      도로명주소 -> 지번 조회키 변환 (건축물대장 조회용, 카카오 로컬 API)
pipeline/fetch_building_ledger.py  건축물대장 표제부 수집 (공공데이터포털 건축HUB API)
pipeline/building_ledger.py    건축물대장 조회키/동 매칭 공통 로직 (--test 로 자체 검사)
pipeline/build_dataset.py      국가DB + 주소 + 지오코딩 캐시 + 관리대수(원본, 프로젝트번호) + 건축물대장을 합쳐 지도 데이터 생성
pipeline/cache/                지오코딩/지번/건축물대장 진행상황 캐시 (git에는 안 올라감)
docs/building-ledger.md        건축물대장 조인 설계 - 조인 키, API 정보, 호출량, 남은 확인사항
config.local.json              카카오 REST API 키 + 공공데이터포털 인증키 (git에는 안 올라감, .example 참고해서 직접 생성)
국가DB_260805.xlsb 등 원본 파일   .gitignore 처리되어 있어 커밋되지 않음 (로컬에만 보관)
```

## 처음 설정

1. [developers.kakao.com](https://developers.kakao.com) 에서 앱 생성 → JavaScript 키, REST API 키 발급
2. `config.local.json.example` 을 복사해서 `config.local.json` 생성 후 키 입력
   - `kakao_rest_api_key`: 카카오 REST API 키
   - `data_go_kr_service_key`: [공공데이터포털](https://www.data.go.kr/data/15134735/openapi.do)
     "국토교통부_건축HUB_건축물대장정보 서비스" 활용신청 후 마이페이지의 **일반 인증키(Decoding)**
3. `index.html` 의 `KAKAO_JS_KEY` 부분을 발급받은 JavaScript 키로 교체
4. 카카오 개발자 콘솔 "플랫폼" 메뉴에 사용할 도메인 등록 (로컬 테스트용 + 이후 GitHub Pages 주소)
5. `pip install pandas pyxlsb requests openpyxl` (필요한 파이썬 패키지)

## 실행 순서

```bash
# 1. 주소 지오코딩 (하루 한도 약 9만 건, 89만 건 전체는 며칠에 걸쳐 반복 실행)
python pipeline/geocode.py

# 2. 지금까지 지오코딩된 만큼 지도 데이터 생성
python pipeline/build_dataset.py

# 3. index.html을 브라우저로 열어서 확인 (또는 로컬 서버로 실행)
```

건축물대장(건물용도) 보강까지 포함하면:

```bash
python pipeline/geocode.py                                    # 1. 위경도
python pipeline/resolve_jibun.py                              # 2. 도로명주소 -> 지번
python pipeline/fetch_building_ledger.py --self-test          # 3. 인증키 확인 (1건 호출)
python pipeline/fetch_building_ledger.py --region "서울특별시 강남구"   # 4. 대장 수집 (파일럿)
python pipeline/build_dataset.py                              # 5. 지도 데이터 생성
```

2~4번은 아직 안 돌려도 된다. 건축물대장 캐시가 없으면 `build_dataset.py`는 기존과 똑같이 동작한다.
자세한 내용(조인 키, 호출량, 남은 확인사항)은 `docs/building-ledger.md` 참고.

지오코딩이 다 끝나지 않아도 `build_dataset.py`를 실행하면 그때까지 완료된 데이터만 반영되므로,
매일 `geocode.py` → `build_dataset.py` 순서로 실행하면 지도가 점점 채워집니다.

## 화면 기능

- 좌측 상단 검색창: 승강기번호(7자리) / 현장명(건물명) / 프로젝트번호(6자리, 현대엘리베이터 직영 건만 존재) 검색
- 제조업체 / 유지보수업체 표시 on-off: 오티스·현대·티케이(티센)·미쓰비시·기타 5개 그룹으로 묶어서 토글
- 우측 미니 대시보드: 현재 필터링된 데이터의 유지보수업체 그룹별 건수 집계
- 상세정보 패널 상단: 건축물대장 기준 건물용도 / 사용승인일 / 연면적 / 층수 / 대장 승강기 수
  (국가DB의 `BULDPRPOS`는 "제2종근린생활시설-기타"처럼 뭉뚱그려진 값이 25만 건이 넘어서
  건축물대장 주용도+기타용도로 실제 용도를 보강함. 대장 수집 전에는 표시되지 않음)
- 프로젝트번호는 `2026.07월 관리대수(원본).xlsb`(현대엘리베이터 자체 관리대수, "download" 시트의 `승강기번호`/`원PJT` 컬럼)와 `ELEVATORNO`로 조인해서 채움 — 타사 관리 건은 null
