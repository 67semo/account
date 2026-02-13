import pandas as pd
import os
import glob

def print_insurance_notice():
    # 1. 다운로드 폴더 경로 설정 (Windows 사용자 홈 기준)
    user_home = os.path.expanduser('~')
    download_dir = os.path.join(user_home, 'Downloads')

    # 2. 파일 검색 패턴 설정 ('2차결정내역통보서'로 시작하는 모든 엑셀 파일)
    # .xls 및 .xlsx 모두 검색하기 위해 확장자에 와일드카드 사용
    search_pattern = os.path.join(download_dir, '2차결정내역통보서*.xls*')
    
    # 해당 패턴의 파일 리스트 가져오기
    files = glob.glob(search_pattern)

    if not files:
        print("❌ 다운로드 폴더에서 '2차결정내역통보서'로 시작하는 파일을 찾을 수 없습니다.")
        return

    # 3. 가장 최근에 생성된 파일 선택
    # os.path.getctime: 파일 생성 시간 기준 정렬
    latest_file = max(files, key=os.path.getctime)
    print(f"📂 읽어올 파일: {latest_file}")

    try:
        # 4. 엑셀 파일 로드
        # '세번째 행이 헤더'이므로 header=2 (0부터 시작하는 인덱스 기준)
        df = pd.read_excel(latest_file, header=2)

        # 5. 필요한 컬럼 정의
        required_cols = ['성명', '당월분_월보험료(원)', '국고지원금액(원)']

        # 데이터프레임에 해당 컬럼들이 모두 존재하는지 확인
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            print(f"❌ 엑셀 파일 내에 다음 컬럼을 찾을 수 없습니다: {missing_cols}")
            print(f"   현재 파일의 컬럼 목록: {df.columns.tolist()}")
            return

        # 6. 데이터 추출 및 출력
        result_df = df[required_cols]
        
        print("\n📊 [추출 결과]")
        print("-" * 50)
        print(result_df)
        print("-" * 50)

    except Exception as e:
        print(f"❌ 파일 처리 중 오류가 발생했습니다: {e}")

if __name__ == "__main__":
    print_insurance_notice()
