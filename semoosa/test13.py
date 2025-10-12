import pandas as pd
import numpy as np

# ----------------------------------------------------
# 0. 예시 데이터프레임 생성 (사용자님의 실제 데이터로 대체하세요)
# ----------------------------------------------------
data = {
    'no': [101, 101, 102, 102, 103, 103, 104, 104, 105, 105, 106, 106],
    '구분': ['수익', '비용', '수익', '비용', '수익', '비용', '수익', '비용', '수익', '비용', '비용', '비용'],
    '계정과목': ['매출', '부가세', '이자수익', '상품', '용역매출', '복리후생비', '잡이익', '부가세', '접대비', '수수료', '부가세', '광고비'],
    '금액': np.random.randint(50, 500, 12) * 10
}
base_df = pd.DataFrame(data)

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

# 저장할 DataFrame과 Sheet 이름 정의
dataframes_to_save = {
    '수익_포함_그룹_전체': df_profit_group,
    '비용_그룹_부가세_포함': df_cost_tax_included,
    '비용_그룹_부가세_미포함': df_cost_tax_excluded
}

try:
    with pd.ExcelWriter(excel_file_name, engine='xlsxwriter') as writer:
        for sheet_name, dataframe in dataframes_to_save.items():
            # 각 DataFrame을 고유한 시트에 저장
            dataframe.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"\n✅ 데이터 처리가 완료되었으며, '{excel_file_name}' 파일에 성공적으로 저장되었습니다.")

except Exception as e:
    print(f"\n❌ 엑셀 저장 중 오류 발생: {e}")