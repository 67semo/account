import os

import pandas as pd
from semoosa import card
from dotenv import load_dotenv

load_dotenv()

data_dir = os.getenv('data_dir')
temp_dir = os.getenv('temp_dir')
organized_file = '카드자료.xlsx'
contact_obj = '거래처.csv'
input_file = os.path.join(data_dir, organized_file)
cont_file = os.path.join(data_dir, contact_obj)

def handling_data(bc_path):

    # 1. 원본 데이터 로드
    df = pd.read_excel(bc_path, sheet_name=0).iloc[:-1]
    df = card.clean_col_data(df)
    card.for_semusa_form(df)
    contractor_df = pd.read_csv(cont_file)
    #print(contractor_df)
    existing_biz_nums = set(contractor_df['사업자등록번호'])

    # 2. 빈 데이터프레임 생성 (최종 결과물)
    #result_df = pd.DataFrame(columns=['날짜', '구분', '계정과목', '적요', '거래처', '차변', '대변', '사업자등록번호'])\
    contractors = []
    collect_rows = []
    print(df)

    # 3. 각 행을 복식부기 처리
    for _, row in df.iterrows():
        # --- 공통 데이터 추출 ---
        date = row['승인일자']
        description = row['가맹점명']
        print(row['가맹점명'])
        card_nm = row['결제계좌은행명'][:2] + str(row['카드번호'][-4:])
        card_desc = row['결제계좌은행명']

        # 거래처 조회및 신규등록
        biz_num = row['가맹점사업자번호']
        if biz_num not in existing_biz_nums:  # 거래처에 해당사업자가 없는 경우 추가.
            contractors.append({
                '사업자등록번호': biz_num,
                '상호': row['가맹점명'],
                '주소': row['가맹점주소1'],
                '종목': row['가맹점업종'],
                '구분': '카드'
            })
            existing_biz_nums.add(biz_num)

        #부가세 해당여부
        if row['회계코드'] in (card.filltering_lst) and row['부가세'] != 0 :
            # --- 1행: 비용 발생 (차변) ---
            row1 = {
                '날짜': date,
                '구분': '비용',
                '계정과목': row['회계코드'],  # 예: '복리후생비'
                '적요': row['회계코드명'],  # 예: '간식'
                '거래처': description,
                '차변': row['거래금액(원화)'],
                '대변': 0,
                '현장명': row['본부명'],
                'code1': biz_num,
                'name' : row['가맹점명']
            }
            row2 = {
                '날짜': date,
                '구분': '자산',
                '계정과목': '부가세대급금',
                '적요': row['회계코드명'],
                '거래처': description,
                '차변': row['부가세'],
                '대변': 0,
                '현장명': row['본부명'],
                'code1': biz_num,
                'name' : row['가맹점명']
            }
            row3 = {
                '날짜': date,
                '구분': '부채',
                '계정과목': '미지급금',
                '적요': row['회계코드명'],  # 예: '간식'
                '거래처': card_nm,
                '차변': 0,
                '대변': row['승인금액(원화)'],
                '현장명': row['본부명'],
                'code1': row['카드번호'],
                'name' : card_desc,
                'approve': row['승인번호']
            }
            rows = [row1, row2, row3]
        else:
            row1 = {
                '날짜': date,
                '구분': '비용',
                '계정과목': row['회계코드'],  # 예: '복리후생비'
                '적요': row['회계코드명'],  # 예: '간식'
                '거래처': description,
                '차변': row['승인금액(원화)'],
                '대변': 0,
                '현장명': row['본부명'],
                'code1': biz_num,
                'name' : row['가맹점명']
            }
            row2 = {
                '날짜': date,
                '구분': '부채',
                '계정과목': '미지급금',
                '적요': row['회계코드명'],
                '거래처': card_nm,
                '차변': 0,
                '대변': row['승인금액(원화)'],
                '현장명': row['본부명'],
                'code1': row['카드번호'],
                'name' : card_desc,
                'approve': row['승인번호']
            }
            rows = [row1, row2]

        # 결과에 추가
        collect_rows = collect_rows + rows
        #result_df = pd.concat([result_df, pd.DataFrame(rows)], ignore_index=True)

    result_df = pd.DataFrame(collect_rows)
    # 4. 결과 확인
    print(result_df.head(6))  # 첫 3개 거래(6행) 출력

    contractor_df = pd.concat([contractor_df, pd.DataFrame(contractors)], ignore_index=True)

    # 5. CSV로 저장 (선택사항)
    temp_file = os.path.join(temp_dir, 'accounting_entries.csv')
    result_df.to_csv(temp_file, index=False, encoding='utf-8-sig')

    '''
    for row_tuple in df.itertuples():
        print(row_tuple)
    df.to_csv('temp.csv', index=False, encoding='utf-8-sig')
    '''
    from datetime import date

    contractor_df['등록일자'] = contractor_df['등록일자'].fillna(date.today())
    contractor_df.to_csv(cont_file, index=False, encoding='utf-8-sig')
    return True

if __name__ == '__main__':
    handling_data(input_file)

