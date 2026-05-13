# DocMine

HWP/HWPX/PDF 문서를 드라이브 전체에서 찾아 본문을 파싱한 뒤 MariaDB에 적재하고, GUI로 키워드 검색하는 파이프라인 도구입니다.

현재 설치된 버전은 `pip show docmine` 또는 `python -c "import docmine; print(docmine.__version__)"` 로 확인할 수 있습니다.

> 0.3.6 부터 프로젝트 이름이 `hwpmine` → `docmine` 으로 변경되었습니다. 기존 `HWPMINE_ENV` 환경변수와 `%APPDATA%\hwpmine\.env` 경로는 이번 버전까지 호환으로 인식하며, 다음 마이너 버전에서 제거됩니다.

## 파이프라인

```
Step 1  scanner        드라이브 스캔 → CSV 추출 (HWP/PDF 별도)
Step 2  inserter       HWP CSV → COM 파싱 → MariaDB 적재
        pdf_inserter   PDF CSV → PyMuPDF 병렬 파싱 → MariaDB 적재
Step 3  search_gui     DB 기반 키워드 검색 GUI (HWP+PDF 통합)
Step 4  extractor_gui  HWP/HWPX → TXT 변환 GUI
```

## 설치

**사전 준비** — Python 3.10 이상이 설치되어 있고 PATH에 등록되어 있어야 합니다. ([python.org](https://www.python.org) 설치 시 *Add Python to PATH* 체크)

### 권장 — `install.bat` 사용 (Windows)

배포 zip을 풀면 `install.bat`, `docmine.exe`, `.env.example`, `README.md` 가 같은 폴더에 있습니다. `install.bat`을 더블클릭하면 `%LOCALAPPDATA%\Programs\docmine\` 에 복사하고 사용자 PATH 에 등록합니다.

### 수동 설치 (wheel)

```bash
pip install docmine-0.3.6-py3-none-any.whl
```

업그레이드 시:

```bash
pip install --upgrade docmine-0.4.0-py3-none-any.whl
```

pip이 이전 버전 파일을 자동으로 정리한 뒤 새 버전을 설치하므로, 구조가 바뀌어도 "고아 파일"이 남지 않습니다. 의존성(`pymysql`, `python-dotenv`, `pymupdf`, Windows의 경우 `pywin32`)은 pip이 자동 설치합니다.

## 환경 설정 — `.env`

설치 후 DB 접속 정보 등을 담은 `.env` 파일을 아래 중 한 곳에 두세요. 위에서부터 우선 적용됩니다.

1. 환경변수 `DOCMINE_ENV` (구 `HWPMINE_ENV` 도 폴백) 가 가리키는 경로
2. **프로그램을 실행하는 현재 작업 디렉터리(CWD)의 `.env`** — 가장 일반적
3. Windows: `%APPDATA%\docmine\.env` (구 `hwpmine` 경로도 폴백)
4. 기타 플랫폼: `~/.config/docmine/.env` (구 `hwpmine` 경로도 폴백)

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
PDF_CSV_FILE=pdf_file_list.csv

COMMIT_EVERY=50
COM_RESTART=500
PARSE_TIMEOUT=60
PDF_WORKERS=0          # 0 또는 미지정 시 os.cpu_count() 자동
```

배포물에 동봉된 `.env.example`을 복사해서 작성하는 것을 권장합니다. `.env`는 패키지에 포함되지 않으므로 업그레이드 시 덮어쓰이지 않습니다.

## 실행

설치 후 어느 디렉터리에서든 단일 런처 `docmine` 명령으로 실행합니다 (pip이 `Python\Scripts\docmine.exe` 런처를 생성).

```bash
docmine            # 대화형 메뉴 (스캔/적재/검색/추출 선택)
docmine g          # 통합 GUI (HWP/PDF 탭 전환)
docmine 1          # Step 1 — HWP 스캐너
docmine 2          # Step 2 — HWP 적재
docmine 3          # Step 3 — 검색 GUI (HWP+PDF 통합)
docmine 4          # Step 4 — HWP 추출 GUI
docmine all        # 1 → 2 → 3 순차 실행 (HWP 파이프라인)
```

PDF 적재는 통합 GUI(g) 의 PDF 탭 또는 다음 모듈 호출로 진입:

```bash
python -m docmine.pdf_inserter
python -m docmine.pdf_inserter --workers 8
```

런처가 인식되지 않는 환경에서는 동일하게 동작하는 모듈 호출을 사용:

```bash
python -m docmine          # = docmine
python -m docmine 3        # = docmine 3
```

## 파일 구조

```
docmine/
├── pyproject.toml           # 패키지 메타데이터 / 의존성
├── README.md
├── .env.example             # 사용자 설정 템플릿
├── run.pyw                  # 콘솔 없이 통합 GUI 실행 (pythonw)
└── src/docmine/
    ├── __init__.py
    ├── __main__.py          # python -m docmine 진입점
    ├── cli.py               # 통합 런처
    ├── config.py            # .env 로드, 전 모듈 공유 설정
    ├── scanner.py           # Step 1 — 파일 스캔
    ├── inserter.py          # Step 2 — HWP 적재 (COM 워커)
    ├── pdf_inserter.py      # Step 2' — PDF 적재 (PyMuPDF 병렬)
    ├── pdf_parser.py        # PDF 텍스트 추출
    ├── search_gui.py        # Step 3 — 통합 검색 GUI
    ├── extractor_gui.py     # Step 4 — HWP→TXT 변환 GUI
    ├── unified_gui.py       # 통합 GUI (4단계 탭)
    ├── drive_picker.py      # 드라이브 선택 다이얼로그
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

PyInstaller 단일 실행파일 빌드 + 배포 zip 패키징을 한 번에:

```bash
build         # build.bat — 빌드 후 docmine_v<버전>.zip 까지 생성
```

`build.bat`은 더블클릭으로도 동작합니다. 두 단계를 직접 돌리고 싶다면:

```bash
pyinstaller docmine.spec --clean --noconfirm   # dist\docmine.exe 생성
python make_release.py                          # docmine_v<버전>.zip 생성
```

버전은 `src/docmine/__init__.py` 의 `__version__` 을 그대로 사용합니다.

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
    parse_status ENUM('success','error','skip','empty') DEFAULT 'success',
    error_msg    VARCHAR(1000),
    parsed_at    DATETIME       DEFAULT CURRENT_TIMESTAMP,
    UNIQUE KEY   uq_file (directory(500), filename(255))
);
```

- `parse_status='empty'` 는 스캔본/이미지 PDF 처럼 본문 텍스트가 비어 있는 정상 케이스 — 진짜 에러(`error`) 와 구분됩니다.

## 동작 방식

### Step 2 — HWP 적재 (크래시 격리)

`inserter` 는 HWP COM 파싱을 별도 워커 프로세스에서 실행합니다. HWP 파일이 COM 을 크래시시켜도 워커만 죽고 메인 프로세스는 새 워커를 띄워 계속 진행합니다.

- `PARSE_TIMEOUT` 초 내에 응답이 없으면 워커 크래시로 판단
- 워커 재시작 후 해당 파일은 에러로 기록 (`hwp_parse_errors.csv`)
- `COM_RESTART` 건마다 COM 인스턴스를 재생성하여 메모리 누수 방지
- 한/글 COM 은 멀티 인스턴스 자동화에 적합하지 않아 단일 워커 직렬 처리

### Step 2' — PDF 적재 (멀티프로세스 병렬)

`pdf_inserter` 는 PyMuPDF 기반 파서를 `multiprocessing.Pool` 로 병렬 실행합니다.

- 워커 수는 `os.cpu_count()` 로 자동 결정 (env `PDF_WORKERS` / CLI `--workers` 로 override)
- 메인 프로세스는 결과를 받아 DB INSERT 만 수행 (단일 커넥션 직렬)
- 스캔본/이미지 PDF 는 `parse_status='empty'` 로 명시적 분리 (에러로 묶이지 않음)

### Step 3 — 통합 검색 GUI

- 공백으로 구분된 키워드 AND/OR/전체문자열 검색
- 검색 대상: 제목+본문 / 제목만 / 본문만 선택 가능
- ID 범위 필터로 특정 ID 이상/이하 범위만 재조회 가능
- 결과 200건씩 페이지 로드, 더보기/전체 조회 지원
- 셀 호버 시 전체 내용 툴팁 표시
- 더블클릭 또는 [파일 열기] 버튼으로 파일 직접 열기
- Del 키로 제외/완전 삭제, Shift+방향키로 범위 선택
- HWP/HWPX/PDF 가 같은 테이블에 적재되어 한 번의 검색으로 모두 조회
