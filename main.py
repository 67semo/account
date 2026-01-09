import sys, os
import time  # 작업 시간 시뮬레이션을 위해 추가
from PyQt5.QtWidgets import QApplication, QMainWindow, QAction, QFileDialog, QMessageBox, QVBoxLayout, QWidget, QLabel
from PyQt5.QtCore import Qt
from semoosa import card11, contactor, card2


class MainWindow(QMainWindow):
    """
    메인 애플리케이션 창 클래스입니다.
    """

    def __init__(self):
        super().__init__()
        self.setWindowTitle("업무 자동화 도구")
        self.setGeometry(100, 100, 600, 400)
        self.setStyleSheet("font-family: 'Malgun Gothic';")

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout(self.central_widget)
        self.info_label = QLabel("메뉴를 선택하여 작업을 시작하세요.", self.central_widget)
        self.info_label.setAlignment(Qt.AlignCenter)
        self.info_label.setStyleSheet("font-size: 16px; font-weight: bold; padding: 20px;")
        self.layout.addWidget(self.info_label)

        self.create_menus()

    def create_menus(self):
        """
        메뉴바와 메뉴 항목들을 생성하고 연결합니다.
        """
        main_menu = self.menuBar()
        card_menu = main_menu.addMenu("카드")
        bank_menu = main_menu.addMenu("은행")
        contact_menu = main_menu.addMenu("거래처")

        # 비씨카드 승인내역 다운받은 압축롸일 기초작업
        account_reco_action = QAction("초기화", self)
        account_reco_action.triggered.connect(self.card_initial_prgress)
        card_menu.addAction(account_reco_action)

        # 카드정리 메뉴 항목
        card_reco_action = QAction("카드정리", self)
        card_reco_action.triggered.connect(self.card_reconciliation)
        card_menu.addAction(card_reco_action)

        # 보고서작성 메뉴 항목
        report_gen_action = QAction("보고서작성", self)
        report_gen_action.triggered.connect(self.generate_report)
        card_menu.addAction(report_gen_action)

        # temp 거래처집계
        gether_cont_action = QAction("매입거래처정리", self)
        gether_cont_action.triggered.connect(self.gatner_contactors)
        contact_menu.addAction(gether_cont_action)

        # temp 매출거래처집계
        gether_cust_action = QAction("매출거래처정리", self)
        gether_cust_action.triggered.connect(self.gatner_customers)
        contact_menu.addAction(gether_cust_action)

    def card_reconciliation(self):
        """
        '카드정리' 메뉴를 위한 슬롯 함수입니다.
        """
        file_path, _ = QFileDialog.getOpenFileName(self, "카드자료 엑셀 파일 선택", "", "Excel Files (*.xls *.xlsx)")
        if file_path:
            self.info_label.setText(f"카드 엑셀 파일 '{file_path.split('/')[-1]}'을(를) 선택했습니다.")
            # 작업 함수를 호출하고 결과를 받습니다.
            success, message = self._process_card_file(file_path)
            self._show_result_message(success, message)

    def card_initial_prgress(self):
        # 1. 우선 조회할 디렉토리를 '다운로드' 폴더로 지정합니다.
        #    사용자 홈 디렉토리를 기반으로 경로를 구성하는 것이 운영체제에 독립적이고 좋습니다.
        download_path = os.path.join(os.path.expanduser('~'), 'Downloads')

        # 2. 파일 필터를 설정합니다.
        #    "압축 파일 (*.zip)"과 같이 사용자에게 친숙한 이름과 확장자를 지정합니다.
        #    여러 필터를 추가할 수도 있습니다. 예: "압축 파일 (*.zip *.rar *.7z)"
        file_filter = "압축 파일 (*.zip)"

        fname = QFileDialog.getOpenFileName(self, '압축 파일 선택', download_path, file_filter)[0]

        if fname[0]:
           card2.card_approval_init(fname, '12')

        else: 
            print("취소됨")


    def account_reconciliation(self):
        """
        '계좌정리' 메뉴를 위한 슬롯 함수입니다.
        """
        file_path, _ = QFileDialog.getOpenFileName(self, "계좌 엑셀 파일 선택", "", "Excel Files (*.xls *.xlsx)")
        if file_path:
            self.info_label.setText(f"계좌 엑셀 파일 '{file_path.split('/')[-1]}'을(를) 선택했습니다.")
            # 작업 함수를 호출하고 결과를 받습니다.
            success, message = self._process_account_file(file_path)
            self._show_result_message(success, message)

    def gatner_contactors(self):
        '''
        "거래처" 수집을 위한 슬롯함수
        '''
        file_path, _ = QFileDialog.getOpenFileName(self, "매입 세금계산서", "", "Excel Files (*.xls *.xlsx)")
        if file_path:
            self.info_label.setText(f"홈텍스 엑셀 파일 '{file_path.split('/')[-1]}'을(를) 선택했습니다.")
            # 작업 함수를 호출하고 결과를 받습니다.
            success, message = self._process_gather_contactors(file_path)
            self._show_result_message(success, message)

    def gatner_customers(self):
        '''
        "거래처" 수집을 위한 슬롯함수
        '''
        file_path, _ = QFileDialog.getOpenFileName(self, "매출 세금계산서", "", "Excel Files (*.xls *.xlsx)")
        if file_path:
            self.info_label.setText(f"홈텍스 엑셀 파일 '{file_path.split('/')[-1]}'을(를) 선택했습니다.")
            # 작업 함수를 호출하고 결과를 받습니다.
            success, message = self._process_gather_customers(file_path)
            self._show_result_message(success, message)


    def generate_report(self):
        """
        '보고서작성' 메뉴를 위한 슬롯 함수입니다.
        """
        file_path, _ = QFileDialog.getOpenFileName(self, "장부 엑셀 파일 선택", "", "Excel Files (*.xls *.xlsx)")
        if file_path:
            self.info_label.setText(f"장부 엑셀 파일 '{file_path.split('/')[-1]}'을(를) 선택했습니다.")
            # 작업 함수를 호출하고 결과를 받습니다.
            success, message = self._generate_report(file_path)
            self._show_result_message(success, message)

    def _show_result_message(self, success, message):
        """
        작업 결과에 따라 다른 메시지 박스를 표시합니다.
        """
        if success:
            QMessageBox.information(self, "작업 완료", message)
        else:
            QMessageBox.critical(self, "작업 오류", message)

    def _process_card_file(self, file_path):
        print(f"카드정리 작업 시작: {file_path}")
        try:
            card11.handling_data(file_path)
            # 예시: process_card_data(file_path)
            time.sleep(1)  # 작업 시간 시뮬레이션
            # 작업이 성공했다고 가정합니다.
            return True, f"'{file_path.split('/')[-1]}' 파일의 카드정리 작업이 성공적으로 완료되었습니다."
        except Exception as e:
            # 작업 중 에러가 발생했다고 가정합니다.
            print(f"카드정리 작업 중 오류 발생: {e}")
            return False, f"'{file_path.split('/')[-1]}' 파일 처리 중 오류가 발생했습니다: {e}"

    def _process_account_file(self, file_path):
        """
        계좌 파일 처리를 위한 더미 함수입니다.
        실제 계좌 정리 로직을 여기에 구현하고,
        성공 여부(True/False)와 메시지(str)를 반환하세요.
        """
        print(f"계좌정리 작업 시작: {file_path}")
        try:
            # 여기에 사용자의 계좌 정리 루틴을 호출하는 코드를 작성합니다.
            # 예시: process_account_data(file_path)
            time.sleep(1)  # 작업 시간 시뮬레이션
            # 작업이 성공했다고 가정합니다.
            return True, f"'{file_path.split('/')[-1]}' 파일의 계좌정리 작업이 성공적으로 완료되었습니다."
        except Exception as e:
            # 작업 중 에러가 발생했다고 가정합니다.
            print(f"계좌정리 작업 중 오류 발생: {e}")
            return False, f"'{file_path.split('/')[-1]}' 파일 처리 중 오류가 발생했습니다: {e}"

    def _generate_report(self, file_path):
        """
        보고서 생성을 위한 더미 함수입니다.
        실제 보고서 작성 로직을 여기에 구현하고,
        성공 여부(True/False)와 메시지(str)를 반환하세요.
        """
        print(f"거래처 작성 작업 시작: {file_path}")
        try:
            time.sleep(1)  # 작업 시간 시뮬레이션
            # 작업이 실패했다고 가정합니다.
            raise ValueError("장부 데이터에 유효하지 않은 값이 포함되어 있습니다.")
        except Exception as e:
            # 작업 중 에러가 발생했다고 가정합니다.
            print(f"보고서작성 작업 중 오류 발생: {e}")
            return False, f"'{file_path.split('/')[-1]}' 파일을 기반으로 보고서 작성 중 오류가 발생했습니다: {e}"

    def _process_gather_contactors(self, file_path):
        try:
            contactor.get_trading_data(file_path)
            time.sleep(1)
            return True, f"'{file_path.split('/')[-1]}' 파일의 거래처 정리작업이 성공적으로 완료되었습니다."
        except Exception as e:
            print(f"카드정리 작업 중 오류 발생: {e}")
            return False, f"'{file_path.split('/')[-1]}' 파일 처리 중 오류가 발생했습니다: {e}"

    def _process_gather_customers(self, file_path):
        try:
            contactor.get_cust_data(file_path)
            time.sleep(1)
            return True, f"'{file_path.split('/')[-1]}' 파일의 거래처 정리작업이 성공적으로 완료되었습니다."
        except Exception as e:
            print(f"카드정리 작업 중 오류 발생: {e}")
            return False, f"'{file_path.split('/')[-1]}' 파일 처리 중 오류가 발생했습니다: {e}"

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
