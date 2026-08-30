# 승강기 서비스 영업현장 맵핑

전국 승강기 설치 위치를 카카오맵에 찍어 서비스 보수 영업에 활용하기 위한 프로젝트.

## 폴더 구조

```
index.html                     지도 페이지 (카카오맵 + 마커 클러스터링)
data/elevators.json            지도가 읽는 실데이터 (pipeline/build_dataset.py가 생성, git 커밋 대상)
data/elevators.sample.json     실데이터 없을 때 보여줄 샘플 5건
pipeline/geocode.py            주소 -> 위경도 지오코딩 (카카오 로컬 API, 하루 한도만큼씩 실행)
pipeline/build_dataset.py      원본 DB + 지오코딩 캐시를 합쳐 data/elevators.json 생성
pipeline/cache/                지오코딩 진행상황 캐시 (git에는 안 올라감)
config.local.json              카카오 REST API 키 (git에는 안 올라감, .example 참고해서 직접 생성)
국가DB_260805.xlsb 등 원본 파일   .gitignore 처리되어 있어 커밋되지 않음 (로컬에만 보관)
```

## 처음 설정

1. [developers.kakao.com](https://developers.kakao.com) 에서 앱 생성 → JavaScript 키, REST API 키 발급
2. `config.local.json.example` 을 복사해서 `config.local.json` 생성 후 REST API 키 입력
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

지오코딩이 다 끝나지 않아도 `build_dataset.py`를 실행하면 그때까지 완료된 데이터만 반영되므로,
매일 `geocode.py` → `build_dataset.py` 순서로 실행하면 지도가 점점 채워집니다.
