# N쇼핑 순위검색 (Naver Search API)

## How to use

- google sheet 에서 1로 표시된 항목들만 순위 찾아서
- google sheet 에 기록

---

## Update

### 230927 수

- 특정 MID 상품만 찾을 것인지, 스토어 상품 모두 찾을 것인지 선택 (todo)

### 230924 일 PM11

- 시트 기록 입력 받기 (turtle -> pyautogui)
- (실패)pyinstaller - 맥에서는 입력 받는 것까지만 작동

### 230923 토 PM10

- 시트 기록 입력 받기

### 230922 금 AM8

- 스토어에 여러개 상품 순위 찾을 수 있게 함
- datatype: dictionary

### 230919 화 오후5

- git 연동 -> Heroku(Cancel) -> Koyeb 업로드\
  - 1분마다 인스턴스 생성되어 돌아가고 있음 (스케쥴 관리 따로 없음)\
    -> 스케쥴 cron 찾기\
    -> Django Project 로 웹에서 관리하고,

### 230905 화 오후2

- 검색페이지 api의 json은 권한 오류로 사용불가
- 대체 - catalog 상품은 제목으로 찾기
- 구글 시트에 기록

### 230904 월

- 네이버 검색 API 이용하여 순위 찾기\
  - 28페이지 29위까지 검색가능 (1000위부터 100개까지)
- 구글 시트에서 읽기
