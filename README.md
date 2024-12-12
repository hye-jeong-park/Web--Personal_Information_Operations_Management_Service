# Web--Personal_Information_Operations_Management_Service
This is the web service version of the Personal Information Operations Management System.


# 개인정보 처리 웹 서비스 사용 방법 (Visual Studio 환경 기준)

## 1. 사전 준비

### 1-1. Python 설치

1. **Python 다운로드 및 설치**:
    - [Python 공식 사이트](https://www.python.org/downloads/)에서 최신 Python 3.x 버전을 다운로드합니다.
    - 설치 시 "Add Python to PATH" 옵션을 체크해주세요. 그래야 명령 프롬프트에서 바로 Python을 사용할 수 있습니다.
2. 설치 완료 후, 명령 프롬프트(또는 터미널)에서 `python --version` 명령을 입력해 버전을 확인합니다.
정상적으로 버전이 표시되면 Python 설치가 완료된 것입니다.

### 1-2. Visual Studio 설치

1. **Visual Studio 다운로드**:
    - [Visual Studio 다운로드 페이지](https://visualstudio.microsoft.com/ko/)에서 Community 버전을 다운로드합니다. (무료)
2. **설치 시 구성 요소 선택**:
    - Visual Studio 설치 관리자에서 "Python 개발" 워크로드를 선택합니다.
    이를 통해 Visual Studio에서 Python 개발 환경을 쉽게 사용할 수 있습니다.
3. 설치가 끝나면 Visual Studio를 실행합니다.

### 1-3. Selenium 및 Flask 환경 구축

1. 명령 프롬프트(또는 Visual Studio 내 통합 터미널)에서 다음 명령을 입력해 필수 라이브러리를 설치합니다:
    
    ```
    pip install selenium openpyxl flask
    ```
    
2. 이로써 Python 스크립트 실행에 필요한 기본적인 라이브러리가 준비됩니다.

---

## 2. 프로젝트 코드 준비하기

1. 제공된 **프로젝트 폴더**를 원하는 경로에 준비합니다. 예를 들어, `C:\Users\사용자명\MyPythonProject` 경로에 놓았다고 가정하겠습니다.
2. 프로젝트 폴더 구조는 대략 다음과 같습니다:
    
    ```
    MyPythonProject/
      ├─ scripts/
      │   ├─ extraction_script.py
      │   ├─ extraction_in_progress_script.py
      │   ├─ delivery_script.py
      │   └─ ... (필요한 추가 스크립트)
      ├─ templates/
      │   └─ index.html
      ├─ static/
      │   ├─ styles.css
      │   ├─ script.js
      │   └─ loading.gif 등
      ├─ app.py
      └─ README.md (이 파일)
    ```
    
3. Visual Studio에서 프로젝트 열기:
    - Visual Studio 실행 후, 상단 메뉴에서 `파일(File)` → `열기(Open)` → `폴더(Folder)`를 선택합니다.
    - `MyPythonProject` 폴더를 선택하고 "폴더 선택"을 클릭하여 프로젝트를 엽니다.
4. 열람 후, `솔루션 탐색기(Solution Explorer)` 패널에서 `app.py` 및 `scripts` 폴더 내의 파이썬 파일들을 확인할 수 있습니다.

---

## 3. 엑셀 파일 및 워크시트 준비

프로그램은 특정 경로에 있는 엑셀 파일에 데이터를 기록합니다. 예를 들어 `extraction_script.py` 파일 내에 다음과 같은 부분이 있습니다:

```python
EXCEL_FILE = r'C:\Users\PHJ\output\개인정보 운영대장.xlsx'
WORKSHEET_NAME = '개인정보 추출 및 이용 관리'

```

위 경로에 맞춰 엑셀 파일을 준비해야 합니다.

1. **엑셀 파일 생성**:
    - 파일 탐색기를 열어 `C:\Users\PHJ\output\` 경로로 이동합니다. (`PHJ`는 예시 사용자명입니다. 실제 사용자명에 맞추어 수정하거나, 경로를 원하는 대로 변경 가능)
    - `output` 폴더가 없다면 새로 만듭니다.
    - Microsoft Excel을 실행하여 빈 통합 문서를 만든 뒤, `개인정보 운영대장.xlsx` 이름으로 `C:\Users\PHJ\output\` 경로에 저장합니다.
2. **워크시트 생성**:
    - 만든 엑셀 파일을 엽니다.
    - 기본 시트명(`Sheet1`)을 우클릭하여 `개인정보 추출 및 이용 관리`로 변경합니다.
    - 이 시트명을 코드상 `WORKSHEET_NAME` 변수가 사용하는 이름과 동일하게 맞춰야 합니다.
3. **경로 수정 (필요 시)**:
    - 다른 경로에 엑셀 파일을 두고 싶다면, `extraction_script.py`, `extraction_in_progress_script.py`, `delivery_script.py`에 있는 `EXCEL_FILE` 경로를 원하는 경로로 수정합니다.
    - 예: `D:\Data\개인정보 운영대장.xlsx` 에 두고 싶다면,
        
        ```python
        EXCEL_FILE = r'D:\Data\개인정보 운영대장.xlsx'
        WORKSHEET_NAME = '개인정보 추출 및 이용 관리'
        ```
        
    - 워크시트명(`WORKSHEET_NAME`)은 엑셀 내부의 시트명과 정확히 동일해야 합니다.

---

## 4. 크롬 드라이버(ChromeDriver) 준비 (Selenium 크롤링용)

Selenium을 사용하여 웹 페이지에 접속, 크롤링을 하려면 Chrome 브라우저와 해당하는 버전의 ChromeDriver가 필요합니다.

1. [ChromeDriver 다운로드 페이지](https://chromedriver.chromium.org/downloads)로 이동하여, 사용하는 Chrome 버전에 맞는 ChromeDriver를 다운로드합니다.
2. 다운로드 받은 `chromedriver.exe`를 Python 스크립트와 같은 폴더(또는 PATH에 잡히는 디렉토리)에 두세요.
    
    혹은 코드 내에서 `webdriver.Chrome()` 호출 시 `executable_path`를 지정할 수 있습니다. 예:
    
    ```python
    driver = webdriver.Chrome(executable_path='C:\\path\\to\\chromedriver.exe')
    ```
    
    현재 제공된 코드에서는 `initialize_webdriver()` 함수 내에 명시적 경로 지정이 없다면, chromedriver.exe를 파이썬 실행 파일이 인식 가능한 디렉토리에 두는 것을 추천합니다.
    

---

## 5. 프로그램 실행 방법

이 프로그램은 Flask(파이썬 웹 프레임워크)를 사용하여 웹 서비스 형태로 실행됩니다. 즉, 웹 브라우저에서 `http://localhost:5000`에 접속하여 사용할 수 있습니다.

1. **Visual Studio에서 Python 스크립트 실행**:
    - `app.py` 파일을 솔루션 탐색기에서 더블클릭하여 엽니다.
    - 상단 메뉴에서 `디버그(Debug)` → `디버깅하지 않고 시작(Start Without Debugging)`를 선택하거나, 바로 실행 아이콘을 클릭하여 실행할 수 있습니다.
    - 처음 실행 시, Python 환경을 묻는 대화상자가 뜰 수 있으니, 설치한 Python 인터프리터를 선택하세요.
    - 콘솔 창(터미널) 또는 출력(Output) 패널에 "Running on [http://127.0.0.1:5000](http://127.0.0.1:5000/)" 혹은 "Running on [http://localhost:5000](http://localhost:5000/)" 문구가 뜨면 서버가 정상적으로 실행된 것입니다.
2. **웹 브라우저에서 접속하기**:
    - Chrome 또는 Edge와 같은 브라우저를 열고 주소창에 `http://localhost:5000`을 입력 후 엔터를 누릅니다.
    - 로그인 화면(아이디, 비밀번호 입력창)과 버튼들이 있는 페이지가 뜨면 준비 완료입니다.

## 6. 프로그램 사용 방법

1. **아이디/비밀번호 입력**:
    
    프로그램은 특정 사내 시스템(예시)에 접속해 크롤링 과정에서 로그인이 필요할 수 있습니다. 해당 시스템에 접근 가능한 실제 아이디와 비밀번호를 입력합니다.
    
2. **크롤링 옵션 선택**:
    - "전체" 옵션: 모든 게시글을 대상으로 처리.
    - "직접입력" 옵션: 처리할 게시글 수를 직접 숫자로 입력할 수 있습니다.
3. **버튼 클릭**:
    - "개인정보 신청 이력 저장": 개인정보 추출 신청서와 관련된 이력을 엑셀에 저장합니다.
    - "개인정보 추출 및 전달": 개인정보 추출 및 전달 관련 데이터도 엑셀에 추가적으로 기록합니다.
4. **처리 완료 후 결과 확인**:
    - 작업이 완료되면 화면에 메시지가 표시되고, 파일 저장 위치가 나타납니다.
    - `개인정보 운영대장.xlsx` 파일을 열어 `개인정보 추출 및 이용 관리` 시트에 데이터가 정상적으로 들어갔는지 확인합니다.

---

## 7. 문제 해결

- **로그인 실패**: 아이디/비밀번호가 맞는지 재확인하고, 해당 홈페이지에 직접 접속이 가능한지(네트워크 환경, VPN 등) 점검하세요.
- **엑셀 파일 오류**:
    - 엑셀 파일 경로를 잘못 지정한 경우, 코드의 `EXCEL_FILE` 변수를 수정해주세요.
    - 워크시트 이름이 불일치할 경우 시트명을 확인하고 코드의 `WORKSHEET_NAME`을 동일하게 맞추세요.
- **라이브러리 관련 오류**:
    - `pip install selenium openpyxl flask` 명령을 다시 실행해 필요한 라이브러리가 설치되었는지 확인하세요.
- **ChromeDriver 오류**:
    - ChromeDriver의 버전이 Chrome 브라우저 버전과 호환되는지 확인하세요.
    - `chromedriver.exe` 파일의 경로가 올바른지 점검하세요.