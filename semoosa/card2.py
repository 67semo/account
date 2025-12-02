# 카드 다운로드 파일을 초기화

import os, zipfile, io, openpyxl, semoosa.xl_utl as xl
import pandas as pd
from dotenv import load_dotenv
from typing import Dict, Optional

load_dotenv()
card_data_dir = os.getenv('data_dir') + '\\25card'

def read_excel_from_zip(zip_file_path):
    """
    ZIP 파일에서 Excel 97-2003 파일(.xls)을 읽어 데이터프레임으로 반환

    Parameters:
    zip_file_path (str): ZIP 파일 경로

    Returns:
    pandas.DataFrame: 읽어온 엑셀 데이터
    """
    try:
        # ZIP 파일 열기
        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            # ZIP 파일 내의 모든 파일 목록 확인
            file_list = zip_ref.namelist()

            # .xls 파일 찾기 (Excel 97-2003 형식)
            xls_files = [f for f in file_list if f.lower().endswith('.xls')]

            if not xls_files:
                raise FileNotFoundError("ZIP 파일 내에 Excel 97-2003 파일(.xls)을 찾을 수 없습니다.")

            # 첫 번째 .xls 파일 선택 (여러 개 있을 경우 첫 번째 파일 사용)
            excel_file_name = xls_files[0]

            # Excel 파일 읽기
            with zip_ref.open(excel_file_name) as excel_file:
                # pandas로 Excel 파일 읽기
                df = pd.read_excel(excel_file, engine='xlrd')

            return df

    except zipfile.BadZipFile:
        raise ValueError("유효하지 않은 ZIP 파일입니다.")
    except Exception as e:
        raise Exception(f"파일 읽기 중 오류 발생: {str(e)}")


def find_and_read_excel_files(directory_path: str, month: str) -> Optional[pd.DataFrame]:
    """
    지정된 디렉토리에서 mm(개월수).xlsx 형식의 파일을 찾아 데이터프레임으로 반환

    Parameters:
    directory_path (str): 검색할 디렉토리 경로
    month (int): 찾고자 하는 개월수

    Returns:
    pandas.DataFrame or None: 찾은 파일의 데이터프레임, 없으면 None
    """
    try:
        # 파일 패턴 생성 (예: 3개월 -> 3.xlsx 또는 03.xlsx)
        month = int(month)
        patterns = [
            f"{month:02d}.xlsx",  # 2자리 숫자로 포맷 (03, 12 등)
        ]
        print(patterns)
        # 디렉토리 내 모든 파일 검색
        for filename in os.listdir(directory_path):
            file_path = os.path.join(directory_path, filename)

            # 파일이고, 패턴 중 하나와 일치하는지 확인
            if os.path.isfile(file_path) and any(filename == pattern for pattern in patterns):
                #print(f"파일 찾음: {filename}")

                # Excel 파일 읽기
                df = pd.read_excel(file_path)
                #print(file_path, df)
                return df

    except Exception as e:
        print(f"read prev. data _오류 발생: {e}")
        return None


def card_approval_init(zip_pth, month):
    print(zip_pth, '/n', month)
    try:
        # 함수 호출
        df = find_and_read_excel_files(card_data_dir, month)  # 월별정리된 엑셀화일의 데이터프레임
        #print(df)

        df1 = read_excel_from_zip(zip_pth)  # 지난3일간의 승인내역
        
        # 결과 출력
        # print("데이터 읽기 성공!")
        print(f"데이터 형태: {df.shape},{df1.shape}")

        uniq_key = '승인번호'
        df1 = df1[:-1]      # 마지막 요약행 제거
        result_df = df1[~df1[uniq_key].isin(df[uniq_key])].sort_values(['승인일자', '승인시간'])

        if result_df.empty:
            print("신규 승인내역이 없습니다.")
        else:
            print(f"신규 승인내역 {result_df.shape[0]}건이 있습니다.")
            trgt_path = os.path.join(card_data_dir, month + '.xlsx')                     # target path
            print(trgt_path)
        
            xl.add_df_to_excel(trgt_path, result_df)                # 기존 엑셀화일에 신규자료 추가
            print("데이터 추가 성공!")

    except Exception as e:
        print(f"오류 발생: {e}")


# 사용 예제
if __name__ == "__main__":
    # ZIP 파일 경로 지정
    zip_path = "C:/Users/PC/Downloads/승인내역조회_20250929053556.zip"  # 실제 파일 경로로 변경 필요

    card_approval_init(zip_path, '09')
