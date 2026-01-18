# gether contector info.
import pandas as pd
from dotenv import load_dotenv
import os

load_dotenv()
data_dir = os.getenv('data_dir')

# 홈텍스에서 다운받을 자료를 갖고 거래처들을 수집 
def collector_contactor_info():
    sales_file = "C:\\Users\\PC\\Downloads\\매출전자세금계산서목록.xls"
    purchase_file = "C:\\Users\\PC\\Downloads\\매입전자세금계산서목록.xls"
    sheet_nm = "세금계산서"

    df = pd.read_excel(purchase_file, sheet_name=sheet_nm, header=5)
    #print('abd', df.columns)

    required_cols = ['공급자사업자등록번호', '상호', '대표자명', '주소', '공급자 이메일']
    contactor_df1 = df[required_cols].drop_duplicates().reset_index(drop=True)


    df1 = pd.read_excel(sales_file, sheet_name=sheet_nm, header=5)
    print('abd', df.columns)
    required_cols1 = ['공급받는자사업자등록번호', '상호1', '대표자명1', '주소1', '공급받는자 이메일1']
    contactor_df2 = df1[required_cols1].drop_duplicates().reset_index(drop=True)

    column_names = ['사업자등록번호', '상호', '대표자명', '주소', '이메일']
    contactor_df1.columns = column_names
    contactor_df2.columns = column_names

    contactor_df = pd.concat([contactor_df1, contactor_df2], ignore_index=True).drop_duplicates().reset_index(drop=True).sort_values(by='사업자등록번호')
    #print(contactor_df)

    contactor_df.to_excel(os.path.join(data_dir, 'contactor_list.xlsx'), sheet_name='거래처', index=False)

# 비씨카드 자료를 갖고 거래처들을 수집 
def collector_contactor1_info(mon):
    card_dr = data_dir + '\\25card'
    low_file = os.path.join(card_dr, f"{mon}.xlsx")
    target_file = os.path.join(data_dir, 'contactor_list.xlsx')
    print(card_dr)
    df = pd.read_excel(low_file)
    #print('abd', df.columns)

    required_cols = ['가맹점사업자번호', '가맹점명', '가맹점업종', '가맹점주소1', '가맹점전화번호']
    contactor_df = df[required_cols].drop_duplicates().reset_index(drop=True)
    column_names = ['사업자등록번호', '상호', '업종', '주소', '전화번호']
    contactor_df.columns = column_names
    contactor_df = contactor_df.drop_duplicates(subset=['사업자등록번호'], keep='last')

    existing_df = pd.read_excel(target_file, sheet_name='거래처')

    merged_df = merge_contactor_info(existing_df, contactor_df)
    merged_df.to_excel(target_file, sheet_name='거래처', index=False)

def merge_contactor_info(df_old, df_new):
    # '사업자등록번호'를 기준으로 인덱스 설정 (업데이트 및 병합의 기준이 됨)
    df_old_idx = df_old.set_index('사업자등록번호')
    df_new_idx = df_new.set_index('사업자등록번호')

    # 1. 기존 데이터프레임 업데이트
    # df_new의 값이 NaN이 아닌 경우에만 df_old를 덮어씁니다.
    df_old_idx.update(df_new_idx)

    # 2. 새로운 거래처 (사업자등록번호가 기존에 없던 행) 찾기 및 추가
    # df_new의 인덱스 중 df_old의 인덱스에 없는 것만 선택합니다.
    new_accounts_idx = df_new_idx.index.difference(df_old_idx.index)

    # 신규 거래처 데이터프레임 생성
    df_only_new = df_new_idx.loc[new_accounts_idx]

    # 기존 데이터 (업데이트 완료)와 신규 데이터 합치기
    df_result = pd.concat([df_old_idx, df_only_new])

    # '사업자등록번호'를 다시 열로 변환
    df_result = df_result.reset_index()

    return df_result

def fill_business_code():
    contactor_file = os.path.join(data_dir, 'contactor_list.xlsx')
    contactor_df = pd.read_excel(contactor_file, sheet_name='거래처')
    
    voucher_file = os.path.join(data_dir, 'voucher_book.xlsx')
    voucher_df = pd.read_excel(voucher_file, sheet_name='4분기')

    # 'unique_code' 열이 비어있는 행만 처리
    for idx, row in voucher_df[(voucher_df['unique_code'].isna()) | (voucher_df['name'].isna())].iterrows():
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
    """
    주어진 거래 데이터프레임을 그룹별로 순차 처리하여 네 가지 유형의 전표
    (매출전표, 카드매입, 매입전표, 일반전표)를 생성합니다.
    """
    # ------------------------------------------------------------------
    # 1. 결과 데이터프레임 초기화
    # ------------------------------------------------------------------
    sales_slip_cols = ['작성날짜', '상호', '사업자등록번호', '대표자', '품명', '공급가액', '부가세', '전표번호']            # 매출전표 열
    card_purchase_cols = ['카드번호', '승인일자', '합계', '공급가', '부가세', '품명', '사업자번호', '상호', '전표번호']     # 카드매입 열
    purchase_slip_cols = ['작성날짜', '상호', '사업자등록번호', '대표자', '품명', '공급가액', '부가세', '전표번호']          # 매입전표 열  
    general_slip_cols = ['날짜', '상호', '적요', '금액', '전표번호', '증빙']                                      # 일반전표 열 

    sales_slip_list = []
    card_purchase_list = []
    purchase_slip_list = []
    general_slip_list = []

    # ------------------------------------------------------------------
    # 2. 'no'열을 그룹화하여 순차적으로 처리
    # ------------------------------------------------------------------
    grouped = voucher_df.groupby('no')
    for no, group in grouped:
        print(no)
        # VAT 관련 행 찾기 (전표번호 전체 그룹에서 찾음)
        vat_rows_in_group = group[group['계정과목'].str.contains('부가세', na=False)]

        # ------------------------------------------------------------------
        # 2. 수익 그룹 처리 (매출전표)
        # ------------------------------------------------------------------
        revenue_rows = group[group['구분'] == '수익']
        print(revenue_rows)
        if not revenue_rows.empty:
            # 2.1. 제외 조건 확인: '구분'이 '수익'인 행 중 '계정과목'이 '잡이익'인 경우
            is_jabyiik = (revenue_rows['계정과목'] == '잡이익').any()
            # 2.1. 제외 조건 확인: 전체 그룹 행 중 '계정과목'이 '수수료'인 행이 있는 경우
            is_susu_row = (group['계정과목'] == '수수료').any()
            
            if not (is_jabyiik or is_susu_row):
                # 전표에 기입할 주된 수익 정보 행 (가장 첫 번째 수익 행 사용)
                main_revenue_row = revenue_rows.iloc[0]
                
                # '공급가액': '대변' - '차변'``
                supply_amount = main_revenue_row['대변'] - main_revenue_row['차변']
                
                # '부가세': '부가세' 계정과목 행의 '대변' - '차변' (수익의 경우 부가세는 대변에 기록)
                vat_amount = 0
                if not vat_rows_in_group.empty:
                    vat_amount = vat_rows_in_group.iloc[0]['대변'] - vat_rows_in_group.iloc[0]['차변']

                new_slip = {
                    '작성날짜': main_revenue_row['날짜'],
                    '상호': main_revenue_row['name'],
                    '사업자등록번호': main_revenue_row['unique_code'],
                    '대표자': main_revenue_row['대표'],
                    '품명': main_revenue_row['적요'],
                    '공급가액': supply_amount,
                    '부가세': vat_amount,
                    '전표번호': no
                }
                print(new_slip)
                sales_slip_list.append(new_slip)

        # ------------------------------------------------------------------
        # 3. 비용 그룹 처리 - 비용 행이 1개인 경우
        # ------------------------------------------------------------------
        expense_rows = group[group['구분'] == '비용']
        
        if len(expense_rows) == 1:
            main_expense_row = expense_rows.iloc[0]
            # '공급가액' 또는 '공급가': '차변' - '대변' (비용의 경우 차변에 기록)
            expense_amount = main_expense_row['차변'] - main_expense_row['대변']
            
            if not vat_rows_in_group.empty:
                # 3.1. VAT 행이 있는 경우
                vat_row = vat_rows_in_group.iloc[0]
                # '부가세': '부가세' 계정과목 행의 '차변' - '대변' (비용의 경우 부가세는 차변에 기록)
                vat_amount = vat_row['차변'] - vat_row['대변']
                
                # 3.1.1. '미지급금' 계정과목 행이 있고 '거래처'가 '카드'로 시작하는 경우 (카드매입)
                # ERROR FIX: .str.startswith('카드', na=False, case=False) 대신 .str.lower().str.startswith('카드'.lower(), na=False) 사용
                #card_check_row = group[(group['계정과목'] == '미지급금') & (group['거래처'].str.startswith('카드', na=False))]
                #card_check_row = group[(group['계정과목'] == '미지급금') & (group['거래처'].str.startswith('카드', na=False))]
                card_check_row = group[(group['계정과목'] == '미지급금') & (group['거래처'].str.contains(r'\d{4}$', na=False))]
                mi = group[group['계정과목']=='미지급금']
                ca = group[group['거래처'].str.startswith('카드', na=False)]
                print(len(mi), len(ca))
                
                if not card_check_row.empty:
                    card_row = card_check_row.iloc[0]
                    card_purchase_list.append({
                        '카드번호': card_row['unique_code'],
                        '승인일자': card_row['날짜'],
                        # '합계': '미지급금' 행의 '대변' - '차변'
                        '합계': card_row['대변'] - card_row['차변'],
                        '공급가': expense_amount,
                        '부가세': vat_amount,
                        '품명': main_expense_row['적요'],
                        '사업자번호': main_expense_row['unique_code'],
                        '상호': main_expense_row['name'],
                        '전표번호': no
                    })
                else:
                    # 3.1.2. 그 외 (매입전표)
                    purchase_slip_list.append({
                        '작성날짜': main_expense_row['날짜'],
                        '상호': main_expense_row['name'],
                        '사업자등록번호': main_expense_row['unique_code'],
                        '대표자': main_expense_row['대표'],
                        '품명': main_expense_row['적요'],
                        '공급가액': expense_amount,
                        '부가세': vat_amount,
                        '전표번호': no
                    })
            else:
                # 3.2. VAT 행이 없는 경우 (일반전표)
                
                # '증빙' 값 결정
                is_deposit = (group['계정과목'] == '보통예금').any()
                # ERROR FIX: .str.startswith('카드', na=False, case=False) 대신 .str.lower().str.startswith('카드'.lower(), na=False) 사용
                is_card_pay = (group['계정과목'] == '미지급금').any() & (group['거래처'].str.startswith('카드', na=False)).any()
                print(is_card_pay)
                proof = None
                if is_deposit:
                    proof = 7
                elif is_card_pay:
                    proof = 1
                else:
                    proof = 0
                
                general_slip_list.append({
                    '날짜': main_expense_row['날짜'],
                    '상호': main_expense_row['거래처'],
                    '적요': main_expense_row['적요'],
                    '금액': expense_amount,
                    '전표번호': no,
                    '증빙': proof
                })

        # ------------------------------------------------------------------
        # 4. 비용 그룹 처리 - 비용 행이 2개인 경우
        # ------------------------------------------------------------------
        elif len(expense_rows) == 2:
            # '차변' - '대변' 금액 계산
            expense_rows.loc[:, '금액'] = expense_rows['차변'] - expense_rows['대변']
            # 금액을 기준으로 정렬 (큰 값이 공급가액, 작은 값이 일반전표 금액)
            sorted_expenses = expense_rows.sort_values(by='금액', ascending=False).reset_index(drop=True)
            
            # 4.1. 매입전표 추가
            main_row_for_purchase = sorted_expenses.iloc[0]
            supply_amount = main_row_for_purchase['금액']
            
            # '부가세' (VAT 행이 있는 경우)
            vat_amount = 0
            if not vat_rows_in_group.empty:
                # '부가세' 계정과목 행의 '차변' - '대변'
                vat_amount = vat_rows_in_group.iloc[0]['차변'] - vat_rows_in_group.iloc[0]['대변']

            purchase_slip_list.append({
                '작성날짜': main_row_for_purchase['날짜'],
                '상호': main_row_for_purchase['name'],
                '사업자등록번호': main_row_for_purchase['unique_code'],
                '대표자': main_row_for_purchase['대표'],
                '품명': main_row_for_purchase['적요'],
                '공급가액': supply_amount, # 큰 값
                '부가세': vat_amount,
                '전표번호': no
            })

            # 4.2. 일반전표 추가
            other_row_for_general = sorted_expenses.iloc[1]
            general_amount = other_row_for_general['금액'] # 작은 값
            
            # '증빙' 값 결정
            is_deposit = (group['계정과목'] == '보통예금').any()
            proof = 7 if is_deposit else None

            general_slip_list.append({
                '날짜': other_row_for_general['날짜'],
                '상호': other_row_for_general['거래처'],
                '적요': other_row_for_general['적요'],
                '금액': general_amount,
                '전표번호': no,
                '증빙': proof
            })
    
    # 리스트를 최종 데이터프레임으로 변환
    sales_slip_df = pd.DataFrame(sales_slip_list, columns=sales_slip_cols)
    card_purchase_df = pd.DataFrame(card_purchase_list, columns=card_purchase_cols)
    purchase_slip_df = pd.DataFrame(purchase_slip_list, columns=purchase_slip_cols)
    general_slip_df = pd.DataFrame(general_slip_list, columns=general_slip_cols)
    
    return sales_slip_df, card_purchase_df, purchase_slip_df, general_slip_df
    

if __name__ == "__main__":
    from time import sleep
    #collector_contactor_info()
    #collector_contactor1_info('09')    # 비씨카드 9월 자료로 거래처 수집
    #sleep(2)  # 파일 생성 대기

    fill_business_code()

    final_xls = os.path.join(data_dir, 'voucher_book_filled.xlsx')
    df = pd.read_excel(final_xls, sheet_name='Sheet1')
   
    report_df = quaterly_report(df)
    with pd.ExcelWriter(os.path.join(data_dir, '4분기_매출전표.xlsx')) as writer:
        report_df[0].to_excel(writer, sheet_name='4분기_매출전표', index=False)
        report_df[1].to_excel(writer, sheet_name='카드매입', index=False)
        report_df[2].to_excel(writer, sheet_name='매입전표', index=False)
        report_df[3].to_excel(writer, sheet_name='일반전표', index=False)
'''

'''