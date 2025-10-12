# 📌 BC카드 승인내역 → 회계장부 자동화 시스템 설계

## ✅ 전체 프로세스 구상

### 1️⃣ BC카드 승인내역 다운로드
- **방법1: BC카드 웹사이트(기업용)**
  - Selenium 같은 웹 자동화 툴로 로그인 후 CSV/엑셀 다운로드
  - (문제: OTP·보안카드 인증 시 자동화 어려움)
- **방법2: ERP API/제휴 API**
  - 카드사에서 공식 API 제공 시 활용 (BC카드는 기업 API 가능성 낮음)
- **방법3: 관리자 수동 다운로드 → 지정 폴더 업로드**
  - 가장 현실적인 방법. 파일을 `downloads/bc_card.csv` 형태로 저장 후 Python이 처리.

---

### 2️⃣ 데이터 전처리 (필요 칼럼만 남기기)
- Pandas 사용해 **거래일자, 가맹점명, 승인금액, 부가세, 승인번호, 카드번호** 등 필요한 칼럼만 남기고 불필요 칼럼 제거.

---

### 3️⃣ 회계 관리자가 추가 입력할 화면
- **GUI 선택지**
  - ✅ **PyQt5 / PySide6** → 데스크톱용 입력 폼
  - ✅ **Django/Flask 웹페이지** → 브라우저에서 입력 가능
- 입력 내용: 계정과목, 거래처, 프로젝트(현장), 메모 등

---

### 4️⃣ 장부(DB) 저장
- ✅ MariaDB/MySQL/SQLite 연결
- ✅ `TransactionOverview` & `TransactionDetail` 같은 테이블에 insert
- ✅ 중복 승인번호 체크 후 이미 입력된 내역은 패스

---

## ✅ Python 코드 예시

```python
import pandas as pd
from sqlalchemy import create_engine

# 1. 카드 승인내역 로드 (CSV 가정)
df = pd.read_csv("bc_card.csv")

# 2. 필요한 칼럼만 남기기
keep_cols = ["거래일자", "승인번호", "카드번호", "가맹점명", "승인금액"]
df = df[keep_cols]

# 3. 화면에 띄우기 (PyQt5 예시)
# - QTableWidget으로 df 보여주고
# - 계정과목, 거래처, 메모 입력란 추가

# 4. DB 저장
engine = create_engine("mysql+pymysql://root:비번@localhost:3306/accounting")

for _, row in df.iterrows():
    sql = """
    INSERT INTO transaction_overview (date, approval_no, card_no, merchant, amount)
    VALUES (%s, %s, %s, %s, %s)
    """
    engine.execute(sql, (row['거래일자'], row['승인번호'], row['카드번호'], row['가맹점명'], row['승인금액']))
