import pandas as pd
import numpy as np

# 1. 예시 데이터프레임 생성
# 2023년 전체 기간의 날짜를 인덱스로 사용합니다.
date_rng = pd.date_range(start='2025-01-01', end='2025-12-31', freq='D')
df = pd.DataFrame(date_rng, columns=['Date'])
print(df)
df['Value'] = np.random.randint(0, 100, size=(len(date_rng)))
df = df.set_index('Date') # 'Date' 컬럼을 DatetimeIndex로 설정

# 데이터프레임 확인
print("원본 데이터프레임 (일부):")
print(df.head())
print("-" * 30)

# 2. 3/4분기 데이터 선택
# 인덱스의 quarter 속성이 3인 행만 선택합니다.
q3_data = df.loc[df.index.quarter == 3]

# 3. 결과 확인
print("3/4분기 데이터 (일부):")
print(q3_data.head())
print(q3_data.tail())

# 3/4분기의 시작일과 끝일 (7월 1일 ~ 9월 30일) 확인
print(f"\nQ3 데이터 시작일: {q3_data.index.min()}")
print(f"Q3 데이터 끝일: {q3_data.index.max()}")