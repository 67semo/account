from dotenv import load_dotenv
import os, datetime, re

load_dotenv()

def get_recent_excel_files(target_directory):
    """
    지정된 디렉터리에서 1시간 이내에 수정된 모든 엑셀 파일의
    파일명을 리스트로 반환합니다.

    Args:
        target_directory (str): 파일을 탐색할 디렉터리 경로.

    Returns:
        list: 조건에 맞는 엑셀 파일명의 리스트.
    """
    # 1. 현재 시간과 8시간 전 시간 계산
    now = datetime.datetime.now()
    one_hour_ago = now - datetime.timedelta(hours=20)

    # 결과를 담을 리스트
    recent_excel_files = []

    # 2. 지정된 디렉터리 유효성 확인
    if not os.path.isdir(target_directory):
        print(f"오류: '{target_directory}' 디렉터리가 존재하지 않습니다.")
        return []

    # 3. 디렉터리 내의 모든 파일 탐색
    for filename in os.listdir(target_directory):
        file_path = os.path.join(target_directory, filename)

        # 파일인지 확인하고 엑셀 파일 확장자(.xlsx, .xls)인지 검사
        if os.path.isfile(file_path) and (filename.endswith('.xlsx') or filename.endswith('.xls')):

            # 파일의 최종 수정 시간(modification time) 가져오기
            mod_time = os.path.getmtime(file_path)

            # Unix 타임스탬프를 datetime 객체로 변환
            mod_datetime = datetime.datetime.fromtimestamp(mod_time)

            # 4. 수정 시간이 8시간 이내인지 확인
            if mod_datetime > one_hour_ago:
                recent_excel_files.append(filename)

    return recent_excel_files


def identify_bank_files(filenames):
    """
    Args:
        filenames (list): 탐색된 파일리스트

    Returns:
        dict: 식별된 엑셀 파일과 해당 은행명이 포함된 딕셔너리.
    """
    #print(filenames)
    # 결과를 담을 딕셔너리
    bank_files_dict = {}

    for filename in filenames:
        # 파일명 규칙에 따라 은행 식별

        # 1. 신한은행 (정확히 일치)
        if filename == 'grid_exceldata.xlsx':
            bank_files_dict['신한은행'] = filename

        # 2. 우리은행
        elif '우리은행 거래내역조회' in filename and filename.endswith('.xlsx'):
            bank_files_dict['우리은행'] = filename

        # 3. 기업은행
        elif '거래내역조회_입출식' in filename and '예금' in filename and filename.endswith('.xlsx'):
            bank_files_dict['기업은행'] = filename

        # 4. 국민은행 (정규표현식 사용)
        # 8자리숫자_14자리숫자_6자리숫자.xls
        elif re.match(r'^\d{8}_\d{14}_\d{6}\.xls$', filename):
            bank_files_dict['국민은행'] = filename

    return bank_files_dict


if __name__ == '__main__':
    data_d = os.getenv('data_dir')
    org_data = os.getenv('original_temp')

    # 파일리스트
    file_list = get_recent_excel_files(org_data)

    # 결과 출력
    if file_list:
        #print(f"'{org_data}' 디렉터리에서 8시간 이내에 수정된 엑셀 파일:")
        file_dic = identify_bank_files(file_list)
    else:
        print(f"'{org_data}' 디렉터리에서 4시간 이내에 수정된 엑셀 파일이 없습니다.")

    # 결과 출력
    if file_dic:
        print("식별된 은행별 엑셀 파일 목록:")
        print(file_dic)
    else:
        print(f"'{org_data}' 디렉터리에서 해당 규칙을 만족하는 엑셀 파일이 없습니다.")