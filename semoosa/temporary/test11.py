import pandas as pd


def filter_by_quarter(df, date_column, quarter_number):
    """
    주어진 데이터프레임에서 특정 분기에 해당하는 행만 필터링합니다.

    Args:
        df (pd.DataFrame): 필터링할 원본 데이터프레임.
        date_column (str): 날짜 정보가 담긴 컬럼 이름 (datetime 형식이어야 함).
        quarter_number (int): 필터링할 분기 (1, 2, 3, 4 중 하나).

    Returns:
        pd.DataFrame: 선택된 분기의 데이터만 포함하는 새로운 데이터프레임.
    """
    if quarter_number not in [1, 2, 3, 4]:
        raise ValueError("분기는 1, 2, 3, 4 중 하나를 입력해야 합니다.")

    # 날짜 컬럼이 datetime 형식인지 확인하고, 아니면 변환 시도
    if not pd.api.types.is_datetime64_any_dtype(df[date_column]):
        print(f"'{date_column}' 컬럼을 datetime 형식으로 변환합니다.")
        df[date_column] = pd.to_datetime(df[date_column])

    # 분기에 따른 월(month) 리스트 정의 (Q1: 1~3, Q2: 4~6, Q3: 7~9, Q4: 10~12)
    month_map = {
        1: [1, 2, 3],
        2: [4, 5, 6],
        3: [7, 8, 9],
        4: [10, 11, 12]
    }

    # 선택된 분기에 해당하는 월 리스트를 가져옵니다.
    target_months = month_map[quarter_number]

    # .dt.month와 .isin()을 사용하여 필터링 조건을 생성합니다.
    filter_condition = df[date_column].dt.month.isin(target_months)

    # 조건에 맞는 행만 선택하여 반환합니다.
    return df[filter_condition]


# ----------------------------------------------------------------------
## 예시 사용
# 1. 예시 데이터프레임 생성
data = {
    '날짜': ['2023-01-05', '2023-04-10', '2023-08-15', '2023-11-20', '2024-03-01'],
    '항목': ['Q1_A', 'Q2_B', 'Q3_C', 'Q4_D', 'Q1_E'],
    '금액': [100, 200, 300, 400, 500]
}
df = pd.DataFrame(data)

# 2. 원하는 분기 입력 (예: 3분기)
try:
    target_quarter = 3  # 여기에 원하는 분기(1, 2, 3, 4)를 입력하세요.

    # 3. 함수 실행 및 결과 저장
    filtered_df = filter_by_quarter(df.copy(), '날짜', target_quarter)

    print(f"--- 원본 데이터프레임 ---")
    print(df)

    print(
        f"\n--- 선택된 {target_quarter}분기 ({', '.join(map(str, filter_by_quarter(df.copy(), '날짜', target_quarter)['날짜'].dt.month.unique()))}월) 데이터 ---")
    print(filtered_df)

except ValueError as e:
    print(f"❌ 에러 발생: {e}")

# 다른 분기 테스트 (예: 1분기)
print("\n" + "=" * 30)
target_quarter_2 = 1
try:
    filtered_df_q1 = filter_by_quarter(df.copy(), '날짜', target_quarter_2)
    print(f"\n--- 선택된 {target_quarter_2}분기 데이터 ---")
    print(filtered_df_q1)
except ValueError as e:
    print(f"❌ 에러 발생: {e}")