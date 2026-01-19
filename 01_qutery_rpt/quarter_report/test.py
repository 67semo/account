# 원장의 기록을 가져와 해당클래스를 분류하고 이를 엑셀화일로 출력하는 process, 
from dotenv import load_dotenv
import os
import pandas as pd

load_dotenv()
dir_path = os.getenv('data_dir')

# 원장 데이터를 가져와 전표단위로 묶고, 고유번호를 지정한후 반환
def change_to_voucher():
    # load draft data
    fname = '25장부.xlsx'
    book = os.path.join(dir_path, fname)
    print(book)
    sheet_nm = "25년장부"
    rough_df = pd.read_excel(book, sheet_name=sheet_nm, header=3)

    # slicing required col.
    req_cols = ['날짜','구분', '계정과목', '적요', '거래처', '차변', '대변', '현장명', 'unique_code', 'name']
    book_df = rough_df[req_cols]

    cols = ["차변", "대변"]
    book_df[cols] = book_df[cols].replace("",0).fillna(0)

    # 결과를 담을 리스트
    no_list = []
    sum_list = []
    group_no = 0
    running_sum = 0

    # 차변 - 대변 계산하면서 누적합
    for debit, credit in zip(book_df["차변"], book_df["대변"]):
        running_sum += (debit - credit)
        no_list.append(group_no)  # 아직 0이 안 된 구간은 0 표시
        sum_list.append(running_sum)

        if running_sum == 0:
            group_no += 1  # 그룹 번호 증가


    # 새로운 열 추가
    book_df.loc[:,"no"] = no_list
    # 유효기간 데이터 추출(3/4분기)
    req_period_df = book_df[book_df['날짜'].dt.quarter == 3]
    req_period_df['날짜'] = req_period_df['날짜'].dt.strftime('%Y-%m-%d')

    return req_period_df


# 전표단위로 묶인 원장 데이터를 '구분' 기준으로 수익과 비용으로 분류하고, '비용' 그룹을 '부가세' 포함 여부에 따라 나누어 엑셀로 저장
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

    excel_file_name = 'Group_Analysis_Results.xlsx'
    report_path = os.path.join(dir_path, excel_file_name)


    # 저장할 DataFrame과 Sheet 이름 정의
    dataframes_to_save = {
        '수익_포함_그룹_전체': df_profit_group,
        '비용_그룹_부가세_포함': df_cost_tax_included,
        '비용_그룹_부가세_미포함': df_cost_tax_excluded
    }

    try:
        with pd.ExcelWriter(report_path, engine='xlsxwriter') as writer:
            for sheet_name, dataframe in dataframes_to_save.items():
                # 각 DataFrame을 고유한 시트에 저장
                dataframe.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"\n✅ 데이터 처리가 완료되었으며, '{excel_file_name}' 파일에 성공적으로 저장되었습니다.")

    except Exception as e:
        print(f"\n❌ 엑셀 저장 중 오류 발생: {e}")

if __name__ == '__main__':
    req_book = change_to_voucher()
    devide_df(req_book)
    #req_book.to_csv('grouped_book.csv', index=False, encoding='utf-8-sig')     # for checking


