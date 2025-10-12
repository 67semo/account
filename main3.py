import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QWidget, QPushButton, QMessageBox
)
from PyQt5.QtCore import Qt


class MainWindow(QMainWindow):
    """
    행 삭제 기능을 포함한 메인 윈도우 클래스
    """

    def __init__(self):
        super().__init__()
        self.setWindowTitle("GUI 테이블 행 삭제 예제")
        self.setGeometry(100, 100, 600, 400)

        # QTableWidget 생성
        self.table_widget = QTableWidget()
        self.init_table()

        # 삭제 버튼 생성
        delete_button = QPushButton("선택한 행 삭제")
        delete_button.setStyleSheet("padding: 10px; font-size: 14px;")
        delete_button.clicked.connect(self.delete_selected_rows)

        # 레이아웃 설정
        layout = QVBoxLayout()
        layout.addWidget(self.table_widget)
        layout.addWidget(delete_button)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    def init_table(self):
        """
        테이블에 더미 데이터를 채웁니다.
        """
        self.table_widget.setRowCount(5)
        self.table_widget.setColumnCount(3)
        self.table_widget.setHorizontalHeaderLabels(['이름', '나이', '도시'])

        data = [
            ('김철수', 25, '서울'),
            ('박영희', 30, '부산'),
            ('이민지', 28, '대구'),
            ('최현우', 35, '인천'),
            ('정수민', 22, '광주')
        ]

        for i, row_data in enumerate(data):
            for j, item in enumerate(row_data):
                self.table_widget.setItem(i, j, QTableWidgetItem(str(item)))

        self.table_widget.resizeColumnsToContents()

    def delete_selected_rows(self):
        """
        선택된 행들을 삭제하는 함수입니다.
        """
        # 선택된 모든 셀 아이템을 가져옵니다.
        selected_items = self.table_widget.selectedItems()

        if not selected_items:
            QMessageBox.warning(self, "경고", "삭제할 행을 선택해주세요.")
            return

        # 선택된 셀 아이템으로부터 유일한 행 번호들을 추출합니다.
        # set을 사용하면 중복된 행 번호를 제거할 수 있습니다.
        rows_to_delete = {item.row() for item in selected_items}

        # 행 번호를 내림차순으로 정렬합니다.
        # 이 부분이 중요합니다. 낮은 인덱스부터 삭제하면 인덱스가 변경되어 문제가 발생할 수 있습니다.
        # 높은 인덱스부터 삭제해야 정확한 행이 삭제됩니다.
        sorted_rows = sorted(list(rows_to_delete), reverse=True)

        reply = QMessageBox.question(self, "확인",
                                     f"{len(sorted_rows)}개의 행을 삭제하시겠습니까?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            for row in sorted_rows:
                self.table_widget.removeRow(row)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
