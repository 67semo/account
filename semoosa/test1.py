# 카드결재내역리스트를 가져와 사용카드의 가지수를 추출(중복제거)

import pandas as pd
import glob
import os

# 디렉토리 경로
folder_path = r"C:\Users\PC\Documents\python\account\semoosa\data\25card"

# 01.xls ~ 08.xls 파일 경로 리스트
file_list = sorted(glob.glob(os.path.join(folder_path, "0[1-8].xlsx")))
print(file_list)
all_data = []

for file in file_list:
    try:
        df = pd.read_excel(file, dtype=str)  # 문자열로 읽기 (데이터 보존)
        # 필요한 열만 추출
        subset = df[["카드번호", "결제계좌은행명"]].dropna(how="all")
        all_data.append(subset)
    except Exception as e:
        print(f"파일 읽기 오류: {file} → {e}")

# 모든 데이터 합치기
if all_data:
    merged_df = pd.concat(all_data, ignore_index=True)
    # 중복 제거
    merged_df = merged_df.drop_duplicates()
    print(merged_df)
    # 리스트로 변환
    result_list = merged_df.values.tolist()

    print(result_list)
else:
    print("처리할 데이터가 없습니다.")
