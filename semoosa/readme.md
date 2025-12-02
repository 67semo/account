# 인덱스
## main.py
### 1. 개요
1. __init__ : gui 초기화
2. create_menus : 메뉴설정및 등록
3. card_reconciliation : 정리된 카드자료.xlsx화일을 읽어와 25장부.xlsx에 기입
4. card_initial_prgress : 다운받은 압축파일을 풀어 추가입력이 강능한 상태의 해당월.xlsx화일로 저장
5. account_reconciliation : 파일다이얼로그로부터 path를 읽어 _process_account_file을 호출하고 그 결과를 _show_result_message로 인계
6. gatner_contactors : 파일다이얼로그로부터 홈텍스엑셀파일 path를 읽어와 _process_gather_contactors를 호출하고 그 결과를 _show_result_message로 인계
7. gatner_customers : 파일다이얼로그로부터 홈텍스엑셀파일 path를 읽어와 _process_gather_customers를 호출하고 그 결과를 _show_result_message로 인계

### 2. 카드업무
1. gui로 실행되어 지나 아직은 미흡함이 많다.
2. bc카드사에 접속하여 카드승인내역을 회계용도로 다운받는다.
3. main.py의 카드>초기화를 선택하면 다운로드디렉토리의 화일매니저 대화창이 실행. 
4. 여기서 다운받은 집화일을 선택하면 card2.py에서 각달의 엑셀화일에 카드 승인내역을 추가

## card2.py
1. card_approval_init에서 컨트롤.
1. find_and_read_excel_files, read_excel_from_zip
2. find_and_read_excel_files에서 기존에 입력된 자료를 가져옴
3. read_excel_from_zip에서 압축을 해제하고 새로 승인된자료를 가져옴
4. 기존의 입력된 자료의 동일한 승인 번호가 있으면 제외하고 나머지의 새로운 승인건만 남겨놓고, xl_util.py의 add_df_to_excel로 인계
5. 그 함수에서 기존에 입력된 화일, 마지막행 다음행부터 판다스의 ExcleWriter를 사용하여 추가

## card.py
    아래의 순서에 의거 생성된 카드 자료을 읽어와 세무사요청양식에 의거 카드자료중 부가세에 해당되는 것등을 추려 카드매입쉬트에 입력
    중간 수동입력이 어차피 필요함. 차라리 장부에 기입을 확인 조정하고 장부에서 정리해오는것이 낳을듯.

    월별자료를 ./data/25card 에 보관하고 계정과목및 적요를 입력하고, .data/카드자료.xlsx에 복붙한후, 
    실행하면 세무사양식의 카드매입 맨 마직막행부터 입력됨.

## card1.py
    아래의 카드자료를 읽어와 장부에 가기재을 하는 코드 (작성중)
# 순서
1. [BC카드](https://www.bccard.com/app/card/MainActn.do)사에 접속하여 승인내역조회 선택
2. 기간을 선택하고 회계양식으로 검색후 엑셀화일로 자료다운.
3. 지정 데이터 디렉토리에 위치시킴.(현재는 데이터 디렉토리의 '카드자료.xlsx')


