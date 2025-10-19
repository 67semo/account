from dotenv import load_dotenv
#from semoosa import card
import os, glob
import pandas as pd

load_dotenv()
data_dir = os.getenv('data_dir')
elec_invoice_nm = '3분기전자세금계산서.xlsx'
filltering_lst = ['재료비','폐기물처리비', '차량유지비', '소모재료비','소모품비','공구비']

def gen_sales_report():
    book = os.path.join(data_dir, elec_invoice_nm)
    sheet_nm = "매출"
    req_cols = ['작성일자', '공급받는자사업자등록번호', '상호', '대표자명', '공급가액', '세액', '품목명']
    rough_df = pd.read_excel(book, sheet_name=sheet_nm, header=5)
    treated_df = rough_df[req_cols].copy()
    treated_df.rename(columns={'공급받는자사업자등록번호': '사업자등록번호'}, inplace=True)
    col_order = ['작성일자', '상호', '사업자등록번호', '대표자명', '품목명', '공급가액', '세액']
    treated_df = treated_df[col_order]
    pa = os.path.join(data_dir, '3분기_매출보고서.xlsx')
    treated_df.to_excel(pa, index=False)

def gen_purchase_report():
    book = os.path.join(data_dir, elec_invoice_nm)
    sheet_nm = "매입"
    req_cols = ['작성일자', '공급자사업자등록번호', '상호', '대표자명', '공급가액', '세액', '품목명']
    rough_df = pd.read_excel(book, sheet_name=sheet_nm, header=5)
    treated_df = rough_df[req_cols].copy()
    treated_df.rename(columns={'공급자사업자등록번호': '사업자등록번호'}, inplace=True)
    col_order = ['작성일자', '상호', '사업자등록번호', '대표자명', '품목명', '공급가액', '세액']
    treated_df = treated_df[col_order]
    pa = os.path.join(data_dir, '3분기_매입보고서.xlsx')
    treated_df.to_excel(pa, index=False)

def gen_card_report(mons:list):
    card_dir = os.path.join(data_dir, '25card')
    books = []

    for mon in mons:
        candidate = os.path.join(card_dir, f'{mon}.xlsx')
        if os.path.exists(card_dir):
            books.append(candidate)
    
    dfs = []
    for f in books:
        try:
            rough_df = pd.read_excel(f)
            #print(rough_df.tail(5))
            rough_df = rough_df[:-1]
            #print(rough_df.tail(5))
            dfs.append(rough_df)
        except Exception as e:
            print(f'Error reading {f}: {e}')

    if not dfs:
        print("No dataframes to concatenate.")
        return
    
    merged = pd.concat(dfs, ignore_index=True, sort=False)

    merged['승인일자'] = pd.to_datetime(merged['승인일자'], errors='coerce')
    merged = merged.set_index('승인일자').sort_index().reset_index()
    start_date = '2025-07-01'
    end_date = '2025-09-30'
    mask = (merged['승인일자'] >= start_date) & (merged['승인일자'] <= end_date)
    merged = merged.loc[mask]

    req_cols = ['카드번호', '승인일자', '가맹점사업자번호', '가맹점명', '승인금액(원화)', '거래금액(원화)','부가세', '회계코드명', '회계코드']
    treated_df = merged[req_cols].copy()

    vats = []
    nors = []
    for _, row in treated_df.iterrows():
        if row['회계코드'] in filltering_lst and row['부가세'] != 0 :
            rov = {
                '카드번호': row['카드번호'],
                '승인일자': row['승인일자'],
                '합계': row['승인금액(원화)'],
                '공급가': row['거래금액(원화)'],
                '부가세': row['부가세'],
                '품명': row['회계코드명'],
                '사업자번호': row['가맹점사업자번호'],
                '상호': row['가맹점명']
            }
            vats.append(rov)
        else:
            ron = {
                '승인일자': row['승인일자'],
                '합계': row['승인금액(원화)'],
                '품명': row['회계코드명'],
                '상호': row['가맹점명'],
                '증빙': '카드'
            }
            nors.append(ron)

    vat_df = pd.DataFrame(vats)
    nor_df = pd.DataFrame(nors)

    excel_file_name = os.path.join(data_dir, 'card_report.xlsx')
    with pd.ExcelWriter(excel_file_name, engine='xlsxwriter') as writer:
        vat_df.to_excel(writer, sheet_name='카드매입', index=False)
        nor_df.to_excel(writer, sheet_name='일반전표', index=False)

    print(f"두 데이터프레임이 '{excel_file_name}' 파일에 성공적으로 저장되었습니다.")
    
if __name__ == "__main__":
    #gen_purchase_report()
    mons = ['07', '08', '09']
    gen_card_report(mons)
