# HWP Mine

HWP/HWPX 파일을 드라이브 전체에서 찾아 본문을 파싱한 뒤 MariaDB에 적재하고, GUI로 키워드 검색하는 파이프라인 도구입니다.

현재 설치된 버전은 `pip show hwpmine` 또는 `python -c "import hwpmine; print(hwpmine.__version__)"` 로 확인할 수 있습니다.

## 파이프라인

```
Step 1  scanner        드라이브 스캔 → CSV 추출
Step 2  inserter       CSV 읽기 → HWP 본문 파싱 → MariaDB 적재
Step 3  search_gui     DB 기반 키워드 검색 GUI
Step 4  extractor_gui  HWP/HWPX → TXT 변환 GUI
```

## 설치

**사전 준비** — Python 3.10 이상이 설치되어 있고 PATH에 등록되어 있어야 합니다. ([python.org](https://www.python.org) 설치 시 *Add Python to PATH* 체크)

### 권장 — `install.bat` 사용 (Windows)

배포 zip을 풀면 `install.bat`, `hwpmine-<버전>-py3-none-any.whl`, `.env.example`, `README.md` 가 같은 폴더에 있습니다. `install.bat`을 더블클릭하면 wheel 을 자동으로 찾아 설치합니다(업그레이드도 동일 — `--force-reinstall` 로 동작).

### 수동 설치

```bash
pip install hwpmine-0.3.0-py3-none-any.whl
```

업그레이드 시:

```bash
pip install --upgrade hwpmine-0.4.0-py3-none-any.whl
```

pip이 이전 버전 파일을 자동으로 정리한 뒤 새 버전을 설치하므로, 구조가 바뀌어도 "고아 파일"이 남지 않습니다. 의존성(`pymysql`, `python-dotenv`, Windows의 경우 `pywin32`)은 pip이 자동 설치합니다.

## 환경 설정 — `.env`

설치 후 DB 접속 정보 등을 담은 `.env` 파일을 아래 중 한 곳에 두세요. 위에서부터 우선 적용됩니다.

1. 환경변수 `HWPMINE_ENV` 가 가리키는 경로
2. **프로그램을 실행하는 현재 작업 디렉터리(CWD)의 `.env`** — 가장 일반적
3. Windows: `%APPDATA%\hwpmine\.env`
4. 기타 플랫폼: `~/.config/hwpmine/.env`

`.env` 내용 예:

```dotenv
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

배포물에 동봉된 `.env.example`을 복사해서 작성하는 것을 권장합니다. `.env`는 패키지에 포함되지 않으므로 업그레이드 시 덮어쓰이지 않습니다.

## 실행

설치 후 어느 디렉터리에서든 단일 런처 `hwpmine` 명령으로 실행합니다 (pip이 `Python\Scripts\hwpmine.exe` 런처를 생성).

```bash
hwpmine            # 대화형 메뉴 (스캔/적재/검색/추출 선택)
hwpmine 1          # Step 1 — 스캐너
hwpmine 2          # Step 2 — 적재
hwpmine 3          # Step 3 — 검색 GUI
hwpmine 4          # Step 4 — 추출 GUI
hwpmine all        # 1 → 2 → 3 순차 실행
```

런처가 인식되지 않는 환경에서는 동일하게 동작하는 모듈 호출을 사용:

```bash
python -m hwpmine          # = hwpmine
python -m hwpmine 3        # = hwpmine 3
```

## 파일 구조

```
hwpmine/
├── pyproject.toml           # 패키지 메타데이터 / 의존성
├── README.md
├── .env.example             # 사용자 설정 템플릿
└── src/hwpmine/
    ├── __init__.py
    ├── __main__.py          # python -m hwpmine 진입점
    ├── cli.py               # 통합 런처
    ├── config.py            # .env 로드, 전 모듈 공유 설정
    ├── scanner.py           # Step 1
    ├── inserter.py          # Step 2
    ├── search_gui.py        # Step 3
    ├── extractor_gui.py     # Step 4
    └── hwp_parser.py        # HWP/HWPX 파싱 라이브러리
```

## 개발자용 — 빌드 / 개발 설치

개발 중 수정사항을 바로 반영하려면 editable install:

```bash
pip install -e .
```

배포용 wheel 빌드:

```bash
pip install build
python -m build         # dist/ 아래에 *.whl 과 *.tar.gz 생성
```

배포 zip 생성 (위에서 빌드한 wheel + install.bat + .env.example + README.md 를 묶음):

```bash
python make_release.py  # dist/ 의 wheel 을 찾아 hwpmine_v<버전>.zip 생성
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

`inserter`는 HWP COM 파싱을 별도 워커 프로세스에서 실행합니다. HWP 파일이 COM을 크래시시켜도 워커만 죽고 메인 프로세스는 새 워커를 띄워 계속 진행합니다.

- `PARSE_TIMEOUT`초 내에 응답이 없으면 워커 크래시로 판단
- 워커 재시작 후 해당 파일은 에러로 기록 (`hwp_parse_errors.csv`)
- `COM_RESTART`건마다 COM 인스턴스를 재생성하여 메모리 누수 방지

### Step 3 검색 GUI

- 공백으로 구분된 키워드 AND/OR/전체문자열 검색
- 검색 대상: 제목+본문 / 제목만 / 본문만 선택 가능
- ID 범위 필터로 특정 ID 이상/이하 범위만 재조회 가능
- 결과 200건씩 페이지 로드, 더보기/전체 조회 지원
- 셀 호버 시 전체 내용 툴팁 표시
- 더블클릭 또는 [파일 열기] 버튼으로 파일 직접 열기
- Del 키로 제외/완전 삭제, Shift+방향키로 범위 선택
