# HWP Mine

HWP/HWPX 파일을 드라이브 전체에서 찾아 본문을 파싱한 뒤 MariaDB에 적재하고, GUI로 키워드 검색하는 파이프라인 도구입니다.

## Version

v 0.1.3

## 파이프라인

```
Step 1  scanner.py    드라이브 스캔 → CSV 추출
Step 2  inserter.py   CSV 읽기 → HWP 본문 파싱 → MariaDB 적재
Step 3  search_gui.py DB 기반 키워드 검색 GUI
```

## 사전 준비

### 의존 패키지 설치

```bash
pip install pymysql python-dotenv pywin32
```

> `pywin32`은 HWP COM 자동화에 필요합니다. HWP가 설치된 Windows 환경에서만 Step 2가 동작합니다.

### 환경 설정

`.env.example`을 복사해 `.env`를 만들고 값을 채웁니다.

```bash
copy .env.example .env
```

```dotenv
# .env
DB_HOST=127.0.0.1
DB_PORT=3306
DB_USER=root
DB_PASSWORD=your_password_here
DB_NAME=hwp_documents
DB_TABLE=documents

SCAN_DRIVES=C:\,D:\
CSV_FILE=hwp_file_list.csv

COMMIT_EVERY=50
COM_RESTART=500
PARSE_TIMEOUT=60
```

### main.py 배치

HWP 파서 라이브러리(`main.py`)를 이 프로젝트 폴더에 함께 두어야 합니다.  
`inserter.py`가 `ZipDocReader`, `SectionParser`, `HWPXDrmError`, `configure_logging`을 해당 파일에서 import합니다.

## 실행

### 통합 런처 (권장)

```bash
python run.py          # 대화형 메뉴
python run.py 1        # Step 1만 실행
python run.py 2        # Step 2만 실행
python run.py 3        # Step 3만 실행
python run.py all      # 1 → 2 → 3 순차 실행
```

### 단독 실행

각 모듈은 런처 없이 직접 실행할 수 있습니다.

```bash
# Step 1 — 스캐너
python scanner.py
python scanner.py --out my_list.csv
python scanner.py --drives "C:\\" "D:\\" "E:\\"

# Step 2 — 적재
python inserter.py --create-db          # DB/테이블만 생성
python inserter.py                       # CSV 전체 적재
python inserter.py --start 0 --end 1000  # 범위 지정
python inserter.py --csv other_list.csv  # CSV 파일 직접 지정

# Step 3 — GUI
python search_gui.py
```

## 파일 구조

```
hwpmine/
├── .env              # DB 접속 정보 및 경로 설정 (git 제외)
├── .env.example      # 팀 공유용 템플릿
├── config.py         # .env 로드, 전 모듈 공유 설정
├── scanner.py        # Step 1: 드라이브 스캔 → CSV
├── inserter.py       # Step 2: CSV → MariaDB 파싱 적재
├── search_gui.py     # Step 3: GUI 검색기
├── run.py            # 통합 런처
└── main.py           # HWP 파서 라이브러리 (별도 배치)
```

## DB 스키마

```sql
CREATE TABLE documents (
    id           INT AUTO_INCREMENT PRIMARY KEY,
    directory    VARCHAR(1000)  NOT NULL,
    filename     VARCHAR(500)   NOT NULL,
    extension    VARCHAR(10)    NOT NULL,
    file_size    BIGINT         DEFAULT 0,
    file_mtime   VARCHAR(30),
    body_text    LONGTEXT,
    parse_status ENUM('success','error','skip') DEFAULT 'success',
    error_msg    VARCHAR(1000),
    parsed_at    DATETIME       DEFAULT CURRENT_TIMESTAMP,
    UNIQUE KEY   uq_file (directory(500), filename(255))
);
```

## 동작 방식

### Step 2 크래시 격리

`inserter.py`는 HWP COM 파싱을 별도 워커 프로세스에서 실행합니다.  
HWP 파일이 COM을 크래시시켜도 워커만 죽고 메인 프로세스는 새 워커를 띄워 계속 진행합니다.

- `PARSE_TIMEOUT`초 내에 응답이 없으면 워커 크래시로 판단
- 워커 재시작 후 해당 파일은 에러로 기록 (`hwp_parse_errors.csv`)
- `COM_RESTART`건마다 COM 인스턴스를 재생성하여 메모리 누수 방지

### Step 3 검색 GUI

- 공백으로 구분된 키워드 AND 검색
- 검색 대상: 제목+본문 / 제목만 / 본문만 선택 가능
- 결과 200건씩 페이지 로드, 더보기/전체 조회 지원
- 셀 호버 시 전체 내용 툴팁 표시
- 더블클릭 또는 [파일 열기] 버튼으로 파일 직접 열기
