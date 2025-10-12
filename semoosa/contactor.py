import pandas as pd
from dotenv import load_dotenv
import os

load_dotenv()
data_dir = os.getenv('data_dir')
obj_path = os.path.join(data_dir, '거래처.csv')
df =  pd.read_csv(obj_path)
#print(df)
contactors_code = list(df['사업자등록번호'])
#print(contactors_code)

def get_trading_data(path):
    print(path)
    df1 = pd.read_excel(path, skiprows=5)
    needed_cols = ['공급자사업자등록번호', '상호', '대표자명', '주소', '공급자 이메일']
    traders = df1.loc[:, needed_cols].drop_duplicates()
    traders_ordered_code = traders.sort_values(by=needed_cols[0])
    traders_ordered_code.to_csv('abc.csv', encoding='utf-8', index=False)
    #print(traders_ordered_code)
    print(len(traders))

    for _, trad in traders.iterrows():
        if trad['공급자사업자등록번호'] not in contactors_code:
            df.loc[len(df)] = {
                '사업자등록번호' : trad['공급자사업자등록번호'],
                '상호' : trad['상호'],
                '대표' : trad['대표자명'],
                '주소' : trad['주소'],
                '이메일' : trad['공급자 이메일']
            }
            contactors_code.append(trad['공급자사업자등록번호'])
    df.to_csv(obj_path, index=False, encoding='utf-8-sig')

def get_cust_data(path):        # 매출데이터
    print(path)
    df1 = pd.read_excel(path, skiprows=5)
    needed_cols = ['공급받는자사업자등록번호', '상호', '대표자명', '주소', '공급받는자 이메일1']
    traders = df1.loc[:, needed_cols].drop_duplicates()
    traders_ordered_code = traders.sort_values(by=needed_cols[0])
    traders_ordered_code.to_csv('abc.csv', encoding='utf-8', index=False)
    #print(traders_ordered_code)
    print(len(traders))

    for _, trad in traders.iterrows():
        if trad['공급받는자사업자등록번호'] not in contactors_code:
            df.loc[len(df)] = {
                '사업자등록번호' : trad['공급받는자사업자등록번호'],
                '상호' : trad['상호'],
                '대표' : trad['대표자명'],
                '주소' : trad['주소'],
                '이메일' : trad['공급받는자 이메일1']
            }
            contactors_code.append(trad['공급받는자사업자등록번호'])
    df.to_csv(obj_path, index=False, encoding='utf-8-sig')


if __name__ == '__main__':
    file_path = r"C:\Users\PC\Downloads\매입전자세금계산서목록(1~84).xls"
    get_trading_data(file_path)