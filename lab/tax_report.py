import os
import pandas as pd
from dotenv import load_dotenv

# Load environment variables
load_dotenv()
data_dir = os.getenv('data_dir')

# Excel 파일 경로
purchase_file = os.path.join(data_dir, '25_4세금계산서.xls')
sales_file = os.path.join(data_dir, '25_4매출매입.xlsx')

print(f"Purchase file: {purchase_file}")
print(f"Sales file: {sales_file}")

# Excel 파일 불러오기
print("\n세금계산서 매입 시트 로드 중...")
purchase_sheet = pd.read_excel(purchase_file, sheet_name='매입', header=5)
print(f"세금계산서 매입 시트 로드 완료: {len(purchase_sheet)} 행")
print(purchase_sheet.head())

print("\n매입매출 매입전표 시트 로드 중...")
sales_sheet = pd.read_excel(sales_file, sheet_name='매입전표')
print(f"매입매출 매입전표 시트 로드 완료: {len(sales_sheet)} 행")

# 사업자등록번호별로 세액 합산
print("\n사업자등록번호별로 세액 합산 중...")
purchase_summary = purchase_sheet.groupby('공급자사업자등록번호')['세액'].sum().reset_index()
purchase_summary.columns = ['사업자등록번호', '세금계산서_세액']

sales_summary = sales_sheet.groupby('사업자등록번호')['부가세'].sum().reset_index()
sales_summary.columns = ['사업자등록번호', '매입전표_세액']

# 데이터 비교
print("\n데이터 비교 중...")
comparison = pd.merge(purchase_summary, sales_summary, on='사업자등록번호', how='outer', suffixes=('_purchase', '_sales'))

# 차액 계산
comparison['세금계산서_세액'] = comparison['세금계산서_세액'].fillna(0)
comparison['매입전표_세액'] = comparison['매입전표_세액'].fillna(0)
comparison['차액'] = comparison['매입전표_세액'] - comparison['세금계산서_세액']

# 차액이 있는 항목만 필터링
incorrect_entries = comparison[comparison['차액'] != 0].copy()
incorrect_entries = incorrect_entries.sort_values('차액', ascending=False)

print(f"\n비교 완료:")
print(f"- 전체 사업자등록번호: {len(comparison)}")
print(f"- 차액 있는 항목: {len(incorrect_entries)}")

if len(incorrect_entries) > 0:
    print("\n차액 있는 항목:")
    print(incorrect_entries.to_string(index=False))

# CSV 파일로 저장
output_file = os.path.join(data_dir, 'tax_discrepancy_report.csv')
incorrect_entries.to_csv(output_file, index=False, encoding='utf-8-sig')
print(f"\n보고서 저장 완료: {output_file}")
