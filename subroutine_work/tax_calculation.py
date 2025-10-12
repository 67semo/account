import pandas as pd
import os
from dotenv import load_dotenv

load_dotenv()
cont_file = os.getenv('correspandent_path')

def tax_invoice(invo_path):
    org_df = pd.read_excel(invo_path,sheet_name="매입", header=5)
    need_cols = ['발급일자', '공급자사업자등록번호', '상호', '대표자명', '주소','합계금액','공급가액', '세액', '전자세금계산서분류','공급자 이메일', '품목명']
    edit_cols = ['발급일자', '상호', '합계금액', '공급가액', '세액', '품목명']

    df = org_df[need_cols]
    edit_df = df[edit_cols].copy()
    edit_df.loc['현장명'] = None

    return df, edit_df

if __name__ == '__main__':
    path = r"C:\Users\PC\Documents\python\account\semoosa\data\세금계산서.xls"
    aa = tax_invoice(path)
    print(aa)