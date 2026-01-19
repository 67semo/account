import os
import pandas as pd
from dotenv import load_dotenv

def check_approval_data():
    # 1. .env 파일 로드 및 경로 설정
    load_dotenv()
    data_dir = os.getenv('data_dir')

    if not data_dir:
        print("❌ .env 파일에서 'data_dir'을 찾을 수 없습니다.")
        return

    # 파일 경로 생성 (OS에 맞게 경로 결합)
    ledger_file = os.path.join(data_dir, '25장부.xlsx')
    approval_file = os.path.join(data_dir, 'sample', '승인내역.xls')

    # 파일 존재 여부 확인
    if not os.path.exists(ledger_file):
        print(f"❌ 파일을 찾을 수 없습니다: {ledger_file}")
        return
    if not os.path.exists(approval_file):
        print(f"❌ 파일을 찾을 수 없습니다: {approval_file}")
        return

    print("📂 데이터를 불러오는 중입니다...")

    # 2. 데이터프레임 로드
    try:
        df_ledger = pd.read_excel(ledger_file, sheet_name='25년장부', header=3)
        df_approval = pd.read_excel(approval_file)
    except Exception as e:
        print(f"❌ 엑셀 파일 읽기 실패: {e}")
        return

    # ==========================================
    # [설정] 엑셀 컬럼명이 다르면 아래 변수를 수정하세요
    # ==========================================
    col_date = '날짜'       # 25장부의 날짜 컬럼명
    col_date1 = '승인일자'   # 승인내역의 날짜 컬럼명
    col_app_no = '승인번호' # 25장부 및 승인내역의 승인번호 컬럼명 (동일하다고 가정)
    # ==========================================
    print("✅ 데이터 로드 완료.", df_ledger.columns.tolist(), df_approval.columns.tolist())
    # 3. 25장부 데이터 전처리 (날짜 필터링)
    # 날짜 형식 변환
    df_ledger[col_date] = pd.to_datetime(df_ledger[col_date])

    # 기간 설정 (2025-10-01 ~ 2025-12-31)
    start_date = '2025-10-01'
    end_date = '2025-12-31'

    # 조건: 기간 내 + 승인번호가 비어있지 않은(NaN이 아닌) 값
    mask = (
        (df_ledger[col_date] >= start_date) & 
        (df_ledger[col_date] <= end_date) & 
        (df_ledger[col_app_no].notna())
    )
    target_ledger = df_ledger.loc[mask].copy()

    # 승인번호를 문자열(str)로 변환 및 공백 제거 (비교 정확도 향상)
    target_ledger[col_app_no] = target_ledger[col_app_no].astype(str).str.strip()
    
    # 승인내역 파일도 승인번호를 문자열로 변환
    df_approval[col_app_no] = df_approval[col_app_no].astype(str).str.strip()

    print(f"\n🔍 분석 대상 기간: {start_date} ~ {end_date}")
    print(f"📊 대상 장부 건수: {len(target_ledger)}건")
    print("-" * 50)

    # 4. 장부 내 중복값 확인 및 보고
    # duplicated: 중복된 항목을 모두 True로 표시 (keep=False)
    dup_mask = target_ledger.duplicated(subset=[col_app_no], keep=False)
    duplicates = target_ledger[dup_mask]

    if not duplicates.empty:
        print("🚨 [경고] 장부에 중복된 승인번호가 있습니다:")
        # 보기 좋게 출력하기 위해 날짜 포맷 변경
        duplicates_print = duplicates[[col_date, col_app_no]].copy()
        duplicates_print[col_date] = duplicates_print[col_date].dt.strftime('%Y-%m-%d')
        print(duplicates_print.sort_values(by=col_app_no).to_string(index=False))
    else:
        print("✅ 장부 내 중복된 승인번호가 없습니다.")

    print("-" * 50)

    # 5. 비교 분석 (누락 확인)
    # 장부의 승인번호 리스트 (중복 제거 set)
    ledger_set = set(target_ledger[col_app_no])
    
    # 승인내역(원본)의 승인번호 리스트 (중복 제거 set)
    approval_source_set = set(df_approval[col_app_no])

    # 비교 1: 장부에는 있는데 승인내역 파일에 없는 것 (오기입 의심)
    only_in_ledger = ledger_set - approval_source_set
    
    # 비교 2: 승인내역 파일에는 있는데 장부에 없는 것 (누락 의심)
    # (단, 기간 내 데이터인지 확인이 어렵다면 단순 비교만 수행)
    missing_in_ledger = approval_source_set - ledger_set

    # 결과 출력
    print("📋 [비교 결과 리포트]")
    
    if len(missing_in_ledger) > 0:
        print(f"\n❗ 승인내역에는 있으나 장부에 누락된 번호 ({len(missing_in_ledger)}건):")
        print(list(missing_in_ledger))
    else:
        print("\n✅ 승인내역에 있는 모든 번호가 장부에 존재합니다.")

    if len(only_in_ledger) > 0:
        print(f"\n❓ 장부에는 있으나 승인내역 파일에서 찾을 수 없는 번호 ({len(only_in_ledger)}건):")
        print(f"   (오타 혹은 취소된 건인지 확인 필요)")
        print(list(only_in_ledger))
    else:
        print("\n✅ 장부의 모든 승인번호가 승인내역 파일에서 확인되었습니다.")

if __name__ == "__main__":
    check_approval_data()