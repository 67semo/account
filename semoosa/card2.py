# 카드 다운로드 파일을 초기화

import os, zipfile, io, openpyxl
import pandas as pd
from dotenv import load_dotenv
from typing import Dict, Optional

load_dotenv()
card_data_dir = os.getenv('data_dir') + '/25card'

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
                print(file_path, df)
                return df

    except Exception as e:
        print(f"오류 발생: {e}")
        return None


def append_df_to_excel(filename, df, sheet_name='Sheet1', startrow=None,
                       truncate_sheet=False, **to_excel_kwargs):            # 집파일 어쩌구 저쩌구 에러가 지속적을 발생 포기 나중에 다시시도.
    """
    DataFrame을 기존 Excel 파일의 마지막 행에 추가하는 함수.

    Args:
        filename (str): Excel 파일 경로.
        df (pd.DataFrame): 추가할 DataFrame.
        sheet_name (str): 시트 이름.
        startrow (int): 데이터가 시작될 행 번호 (기본값: 마지막 행 다음).
        truncate_sheet (bool): 시트를 덮어쓸지 여부 (기본값: False).
        **to_excel_kwargs: pandas.DataFrame.to_excel의 추가 인자.
    """

    # openpyxl 엔진을 사용하여 ExcelWriter 객체 생성
    writer = pd.ExcelWriter(filename, engine='openpyxl')

    try:
        # 기존 워크북 로드
        writer.book = openpyxl.load_workbook(filename)

        # 특정 시트 선택
        if sheet_name in writer.book.sheetnames:
            writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)

        # 마지막 행 찾기
        if startrow is None and sheet_name in writer.sheets:
            startrow = writer.sheets[sheet_name].max_row

        # 기존 데이터를 덮어쓸 경우
        if truncate_sheet and sheet_name in writer.sheets:
            # 기존 시트 제거 및 새로운 시트 추가 (파일을 덮어쓰는 효과)
            std = writer.book[sheet_name]
            writer.book.remove(std)
            writer.book.create_sheet(sheet_name)

            # DataFrame을 Excel 파일에 추가
            df.to_excel(writer, sheet_name, header=True, index=False, **to_excel_kwargs)

            # 파일 저장
            writer.save()

        else:
            # 새로운 DataFrame을 기존 파일의 마지막 행에 추가
            df.to_excel(writer, sheet_name, startrow=startrow, header=False, index=False, **to_excel_kwargs)

            # 파일 저장
            writer.save()

    except FileNotFoundError:
        # 파일이 존재하지 않으면, 새로운 파일 생성 후 DataFrame 저장
        df.to_excel(writer, sheet_name, header=True, index=False, **to_excel_kwargs)
        writer.save()

def card_approval_init(zip_pth, month):
    print(zip_pth, '/n', month)
    try:
        # 함수 호출
        df = find_and_read_excel_files(card_data_dir, month)  # 월별정리된 엑셀화일의 데이터프레임
        df1 = read_excel_from_zip(zip_pth)  # 지난3일간의 승인내역
        print(df)

        # 결과 출력
        # print("데이터 읽기 성공!")
        print(f"데이터 형태: {df.shape},{df1.shape}")

        uniq_key = '승인번호'
        df1 = df1[:-1]
        result_df = df1[~df1[uniq_key].isin(df[uniq_key])].sort_values(['승인일자', '승인시간'])
        trgt_path = os.path.join(card_data_dir, month + '.csv')
        result_df.to_csv(trgt_path, index=False, encoding='utf-8-sig')

    except Exception as e:
        print(f"오류 발생: {e}")



# 사용 예제
if __name__ == "__main__":
    # ZIP 파일 경로 지정
    zip_path = "C:/Users/PC/Downloads/승인내역조회_20250929053556.zip"  # 실제 파일 경로로 변경 필요

    card_approval_init(zip_path, '09')
