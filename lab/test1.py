# gether contector info.
import pandas as pd
from dotenv import load_dotenv
import os

load_dotenv()
data_dir = os.getenv('data_dir')

def collector_contactor_info():
    sales_file = "C:\\Users\\garam\\Downloads\\매출전자세금계산서목록(1~35).xlsx"
    purchase_file = "C:\\Users\\garam\\Downloads\\매입전자세금계산서목록(1~84).xlsx"
    sheet_nm = "세금계산서"

    df = pd.read_excel(purchase_file, sheet_name=sheet_nm, header=5)
    #print('abd', df.columns)

    required_cols = ['공급자사업자등록번호', '상호1', '대표자명1', '주소1', '공급자 이메일']
    contactor_df1 = df[required_cols].drop_duplicates().reset_index(drop=True)


    df1 = pd.read_excel(sales_file, sheet_name=sheet_nm, header=5)
    print('abd', df.columns)
    required_cols1 = ['공급받는자사업자등록번호', '상호1', '대표자명1', '주소1', '공급받는자 이메일1']
    contactor_df2 = df1[required_cols1].drop_duplicates().reset_index(drop=True)

    contactor_df1.columns = ['사업자등록번호', '상호', '대표자명', '주소', '이메일']
    contactor_df2.columns = ['사업자등록번호', '상호', '대표자명', '주소', '이메일']

    contactor_df = pd.concat([contactor_df1, contactor_df2], ignore_index=True).drop_duplicates().reset_index(drop=True).sort_values(by='사업자등록번호')
    #print(contactor_df)

    contactor_df.to_excel(os.path.join(data_dir, 'contactor_list.xlsx'), index=False)

def fill_business_code():
    contactor_file = os.path.join(data_dir, 'contactor_list.xlsx')
    contactor_df = pd.read_excel(contactor_file, sheet_name='거래처')
    
    voucher_file = os.path.join(data_dir, 'voucher_book.xlsx')
    voucher_df = pd.read_excel(voucher_file, sheet_name='3분기')

    # 'code' 열이 비어있는 행만 처리
    for idx, row in voucher_df[voucher_df['unique_code'].isna()].iterrows():
        search_str = str(row['거래처'])
        # '상호'에 거래처 문자열이 포함된 행 찾기
        matches = contactor_df[contactor_df['상호'].astype(str).str.contains(search_str, na=False, regex=False)]
        if len(matches) == 0:
            voucher_df.at[idx, 'unique_code'] = 'non'
        elif len(matches) == 1:
            voucher_df.at[idx, 'unique_code'] = matches.iloc[0]['사업자등록번호']
            voucher_df.at[idx, 'name'] = matches.iloc[0]['상호']
            voucher_df.at[idx, '대표'] = matches.iloc[0]['대표자명']
        # 여러개면 아무것도 안함(필요시 else 추가)


    # 결과를 엑셀로 저장 (필요시 주석 해제)
    voucher_df.to_excel(os.path.join(data_dir, 'voucher_book_filled.xlsx'), index=False)

def quaterly_report(voucher_df):
     sales_slip = pd.DataFrame(columns=['작성날짜', '상호', '사업자등록번호', '대표자', '품명', '공급가액', '부가세', '전표번호'])
     skipped = []

    # 'no'로 그룹화
     grouped = voucher_df.groupby('no')

     for no, group in grouped:
        # '구분'이 '수익'인 행 찾기
         income_rows = group[group['구분'] == '수익']
         if income_rows.empty:
            continue

        # '계정과목'이 '잡이익'인 행이 있거나, '수수료'인 행이 있으면 무시
         skip_reason = None
         if any(income_rows['계정과목'] == '잡이익'):
            skip_reason = '잡이익'
         elif any(group['계정과목'] == '수수료'):
            skip_reason = '수수료'

         if skip_reason:
            skipped.append({'전표번호': no, '원인계정과목': skip_reason})
            continue

        # '구분'이 '수익'인 첫 번째 행 기준
         income_row = income_rows.iloc[0]

        # '부가세'가 포함된 계정과목 행 찾기
         vat_rows = group[group['계정과목'].astype(str).str.contains('부가세')]
         vat_value = 0
         if not vat_rows.empty:
            vat_value = (vat_rows.iloc[0]['대변'] - vat_rows.iloc[0]['차변'])

        # 공급가액 계산
         supply_value = income_row['대변'] - income_row['차변']

        # sales_slip에 추가
         sales_slip = pd.concat([sales_slip, pd.DataFrame([{
            '작성날짜': income_row['날짜'],
            '상호': income_row['거래처'],
            '사업자등록번호': income_row['unique_code'],
            '대표자': income_row['대표'],
            '품명': income_row['적요'],
            '공급가액': supply_value,
            '부가세': vat_value,
            '전표번호': income_row['no']
         }])], ignore_index=True)

     skipped_df = pd.DataFrame(skipped)
     return sales_slip, skipped_df
    

if __name__ == "__main__":
    #collector_contactor_info()
    fill_business_code()
    final_xls = os.path.join(data_dir, 'voucher_book_filled.xlsx')
    df = pd.read_excel(final_xls, sheet_name='Sheet1')
    report_df, skipped_df = quaterly_report(df)
    with pd.ExcelWriter(os.path.join(data_dir, '3분기_매출전표.xlsx')) as writer:
        report_df.to_excel(writer, sheet_name='3분기_매출전표', index=False)
        skipped_df.to_excel(writer, sheet_name='스킵내역', index=False)

