# 카드 결제내역과 비교하며 거래자 상호 고유코드및 상호등을 원장에 입력

import os
import pandas as pd
from dotenv import load_dotenv
from pathlib import Path
from semoosa import card

load_dotenv()

card_folder = Path(os.getenv('card_data_folder'))   # 카드자료 푤더
card_files = list(card_folder.rglob("*.xlsx"))
cont_file = os.getenv('contact_obj')                # 거래처 자료
#book_path = os.getenv('book')
book_path = './data/25장부temp.xlsx'

#book_df = pd.read_excel(book_path, sheet_name='25년장부', header=3)
book_df = pd.read_excel(book_path, sheet_name='변경원장')
#print(book_df)

#print(card_folder, card_files)
other_rows = []
for file in card_files:
    df = pd.read_excel(file).iloc[:-1]
    #df.to_csv('abc.csv')
    for _, row in df.iterrows():
        date = row['승인일자']
        description = row['가맹점명']
        biz_code = row['가맹점사업자번호']
        card_code = row['카드번호']
        card_bank = row['결제계좌은행명']
        card_name = '신용' + row['카드번호'][-4:]

        price = int(row['승인금액(원화)'].replace(',', ''))
        filterd_df = book_df[(book_df['대변'] == price) & (book_df['날짜'] == row['승인일자']) & (book_df['계정과목'] == '미지급금')]
        print(len(filterd_df), row['승인금액(원화)'] )
        if len(filterd_df)  == 1:
            #print(filterd_df.index)
            book_df.loc[filterd_df.index, 'code1'] = card_code
            book_df.loc[filterd_df.index, 'name'] = card_name
            book_df.loc[filterd_df.index-1, 'code1'] = biz_code
            book_df.loc[filterd_df.index-1, 'name'] = description
            if row['회계코드'] in (card.filltering_lst) and row['부가세'] != 0:
                book_df.loc[filterd_df.index - 2, 'code1'] = biz_code
                book_df.loc[filterd_df.index - 2, 'name'] = description
        else:
            other_rows.append(row)

other_df = pd.DataFrame(other_rows)
book_df.to_excel('temp1.xlsx', sheet_name='변경원장')
other_df.to_excel('temp.xlsx', sheet_name='미분류')

grouped = book_df.groupby('no')

