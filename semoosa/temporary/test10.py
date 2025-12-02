import pandas as pd

# 가정된 원본 데이터프레임 (예시)
# 실제 데이터프레임의 컬럼 이름과 타입을 맞춰주세요.
data = {
    'no': [101, 101, 101, 202, 202, 303, 303, 303, 404, 404, 505],
    '날짜': ['2023-10-01', '2023-10-01', '2023-10-01', '2023-10-02', '2023-10-02', '2023-10-03', '2023-10-03',
           '2023-10-03', '2023-10-04', '2023-10-04', '2023-10-05'],
    '계정과목': ['접대비', '보통예금', '부가세예수금', '여비교통비', '보통예금', '복리후생비', '보통예금', '여비교통비', '접대비', '보통예금', '세금과공과'],
    '구분': ['비용', '지급', '부가세', '비용', '지급', '비용', '지급', '비용', '비용', '지급', '비용'],
    '거래처': ['A업체', 'A업체', 'A업체', 'B업체', 'B업체', 'C업체', 'C업체', 'C업체', 'D업체', 'D업체', 'E업체'],
    '차변': [100000, 0, 10000, 50000, 0, 80000, 0, 30000, 120000, 0, 70000],
    '대변': [0, 110000, 0, 0, 50000, 0, 110000, 0, 0, 120000, 0],
    '적요': ['식사', '이체', '부가세', '출장비', '이체', '간식구매', '이체', '택시비', '선물', '이체', '교육비'],
    '비고': ['비고1', '비고2', '비고3', '비고4', '비고5', '비고6', '비고7', '비고8', '비고9', '비고10', '비고11']
}
df = pd.DataFrame(data)

# '차변'과 '대변'이 숫자인지 확인하고, 혹시 문자열이면 숫자로 변환합니다. (에러 방지)
df['차변'] = pd.to_numeric(df['차변'], errors='coerce').fillna(0)
df['대변'] = pd.to_numeric(df['대변'], errors='coerce').fillna(0)

# --- 1. no 컬럼으로 그룹화 ---
grouped = df.groupby('no')


# --- 2. 각 그룹에서 '부가세' 문자열이 포함되지 않고 '구분'이 '비용'인 행이 있는 그룹만 선택 ---
# '계정과목'에 '부가세'가 포함되지 않고 ('~'는 not의 의미), '구분'이 '비용'인 행이 그룹에 '하나 이상' 있는지 확인
def filter_groups(group):
    # 조건을 만족하는 행이 하나라도 있으면 True 반환
    return ((~group['계정과목'].str.contains('부가세', na=False)) & (group['구분'] == '비용')).any()


filtered_groups = grouped.filter(filter_groups)

# 다시 그룹화 (필터링된 데이터에 대해서만 작업)
filtered_grouped = filtered_groups.groupby('no')


# --- 3. 그 그룹 내에서 '구분' 열이 '비용'인 행이 둘 이상이면 에러를 발생시켜 알려줘. ---
# 사용자 정의 집계 함수 (apply()에 사용)
def process_group(group):
    # '구분'이 '비용'인 행을 필터링합니다.
    cost_rows = group[group['구분'] == '비용']
    cost_count = len(cost_rows)
    group_no = group['no'].iloc[0]  # 그룹 번호 가져오기

    # 에러 조건 확인: '비용'인 행이 2개 이상인 경우
    if cost_count > 1:
        # 에러 발생 (no 값을 포함하여 명확하게 알림)
        raise ValueError(f"에러 발생: no={group_no} 그룹 내에 '구분'이 '비용'인 행이 {cost_count}개 있습니다. (2개 이상)")

    # 에러가 발생하지 않고, '비용' 행이 1개인 경우 (이 경우만 유효하다고 가정)
    if cost_count == 1:
        # 4단계 작업을 위한 데이터 추출 및 구성
        cost_row = cost_rows.iloc[0]

        # '비용' 행을 제외한 나머지 행들
        other_rows = group[group['구분'] != '비용']

        # '비고' 컬럼 생성: 나머지 행들의 계정과목과 거래처를 리스트화
        memo_list = [(row['계정과목'], row['거래처']) for index, row in other_rows.iterrows()]

        # 새로운 행 (Series) 생성
        new_row = pd.Series({
            '날짜': cost_row['날짜'],
            '거래처': cost_row['거래처'],
            '금액': cost_row['차변'] - cost_row['대변'],
            '적요': cost_row['적요'],
            '비고': memo_list
        })
        return pd.DataFrame([new_row])

    # '비용' 행이 0개인 경우 (필터링 단계에서 걸러졌을 가능성이 높지만 안전을 위해 빈 DataFrame 반환)
    return pd.DataFrame()


# --- 4. 새로운 데이터프레임 생성 ---
# try-except 블록으로 에러를 포착하고 처리합니다.
try:
    # apply()를 사용하여 각 그룹에 사용자 정의 함수 적용 및 결과 결합
    new_df_list = filtered_grouped.apply(process_group)

    # 결과를 하나의 데이터프레임으로 정리 (MultiIndex 제거)
    new_df = pd.concat(new_df_list.tolist(), ignore_index=True)
    print("✅ 새로운 데이터프레임이 성공적으로 생성되었습니다:\n")
    print(new_df)

except ValueError as e:
    # 3단계에서 발생한 에러 처리
    print(f"❌ 데이터 처리 중 에러가 발생했습니다: {e}")

# '비용' 행이 0개인 그룹을 걸러내지 않은 경우를 대비하여 `process_group`에서 빈 DataFrame 반환 처리를 했습니다.
# 만약 `process_group`이 '비용' 행이 0개인 그룹에 대해서도 어떤 처리를 원한다면 해당 로직을 수정해야 합니다.