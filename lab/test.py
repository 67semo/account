# 원장의 기록을 가져와 해당클래스를 분류하여 전표번호를 부여하고 이를 엑셀화일로 출력하는 process
from dotenv import load_dotenv
import os
import pandas as pd

load_dotenv()
data_dir = os.getenv('data_dir')

def voucher():
    # load draft data
    book = os.path.join(data_dir, '25장부.xlsx')
    #print(book)
    sheet_nm = "25년장부"
    rough_df = pd.read_excel(book, sheet_name=sheet_nm, header=3)

    # slicing required col.
    req_cols = ['날짜','구분', '계정과목', '적요', '거래처', '차변', '대변', '현장명', 'unique_code', 'name']
    book_df = rough_df[req_cols].copy()

    # 날짜 칼럼을 datetime 형식으로 변환
    book_df['날짜'] = pd.to_datetime(book_df['날짜'], errors='coerce')

    cols = ["차변", "대변"]
    book_df.loc[:, cols] = book_df[cols].replace("",0).fillna(0)    # 공백은 0을, 빈값또한 0으로 대체

    # 결과를 담을 리스트
    no_list = []
    sum_list = []
    group_no = 0
    running_sum = 0

    # 차변 - 대변 계산하면서 누적합
    for debit, credit in zip(book_df["차변"], book_df["대변"]):
        running_sum += (debit - credit)
        no_list.append(group_no)  # 아직 0이 안 된 구간은 0 표시, 전표번호리스트
        sum_list.append(running_sum)

        if running_sum == 0:
            group_no += 1  # 그룹 번호 증가, 현재의 마무리된 전표번호


    # 새로운 열 추가
    book_df["no"] = no_list
    # 유효기간 데이터 추출(3/4분기)
    req_period_df = book_df[book_df['날짜'].dt.quarter == 4].copy()         # 분기입력
    req_period_df['날짜'] = req_period_df['날짜'].dt.strftime('%Y-%m-%d')

    return req_period_df

def devide_df(base_df):

    # ----------------------------------------------------
    # 1 & 2. '구분' 기준으로 '비용' 또는 '수익'을 포함하는 그룹 추출 및 슬라이싱
    # ----------------------------------------------------

    # 'no' 그룹에 '비용' 행이 하나라도 있는지 확인하는 필터
    mask_cost = base_df.groupby('no')['구분'].transform(lambda x: (x == '비용').any())
    # 'no' 그룹에 '수익' 행이 하나라도 있는지 확인하는 필터
    mask_profit = base_df.groupby('no')['구분'].transform(lambda x: (x == '수익').any())

    # 2. '비용' 그룹만 포함하는 DataFrame
    df_cost_group = base_df[mask_cost].copy()  # SettingWithCopyWarning 방지
    # 2. '수익' 그룹만 포함하는 DataFrame
    df_profit_group = base_df[mask_profit].copy()  # SettingWithCopyWarning 방지

    # ----------------------------------------------------
    # 3. '비용' 그룹 DataFrame을 '부가세' 포함 여부에 따라 분할
    # ----------------------------------------------------

    # df_cost_group 내에서 'no' 그룹에 '부가세'가 포함된 행이 하나라도 있는지 확인
    mask_tax_in_cost_group = df_cost_group.groupby('no')['계정과목'].transform(
        lambda x: x.str.contains('부가세', na=False).any()
    )

    # 3. 최종 결과 1: '비용' 그룹 중 '부가세' 포함 그룹
    df_cost_tax_included = df_cost_group[mask_tax_in_cost_group]

    # 3. 최종 결과 2: '비용' 그룹 중 '부가세' 미포함 그룹
    df_cost_tax_excluded = df_cost_group[~mask_tax_in_cost_group]

    # ----------------------------------------------------
    # 4. 세 개의 DataFrame을 하나의 엑셀 파일의 다른 시트에 저장
    # ----------------------------------------------------

    excel_file = os.path.join(data_dir,'Group_Analysis_Results.xlsx')

    # 저장할 DataFrame과 Sheet 이름 정의
    dataframes_to_save = {
        '수익_포함_그룹_전체': df_profit_group,
        '비용_그룹_부가세_포함': df_cost_tax_included,
        '비용_그룹_부가세_미포함': df_cost_tax_excluded
    }
    save_to_excel('Group_Analysis_Results.xlsx', dataframes_to_save)

def save_to_excel(file_name, sheet_dict):
    excel_path = os.path.join(data_dir, file_name)
    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
        for sheet_name, df in sheet_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"DataFrame saved to {excel_path} in sheet '{sheet_name}'")

if __name__ == '__main__':
    req_book = voucher()        # 장부내용을 전표화(전표번호부여)
    print(req_book.head())
    save_to_excel('voucher_book.xlsx', {'4분기': req_book})    # 저장
    #devide_df(req_book)    # analysis and divide
    #req_book.to_csv('grouped_book.csv', index=False, encoding='utf-8-sig')     # for checking
