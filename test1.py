import sys
import pandas as pd
import traceback
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QWidget, QPushButton, QTabWidget
)
from PyQt5.QtCore import Qt


class TableViewerWithTabs(QMainWindow):
    def __init__(self, df):
        super().__init__()
        self.setWindowTitle('Pandas DataFrame Viewer')
        self.setGeometry(100, 100, 900, 700)

        # 원본 DataFrame을 저장하여 수정된 값과 비교할 수 있게 함
        self.original_df = df.copy()

        # '금액' 열을 숫자 타입으로 변환 (천 단위 포맷팅을 위해)
        self.df = df
        self.df['금액'] = pd.to_numeric(self.df['금액'], errors='coerce').fillna(0).astype(int)

        # 메인 위젯 및 레이아웃 설정
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)

        # 탭 위젯 생성
        self.tab_widget = QTabWidget()
        self.main_layout.addWidget(self.tab_widget)

        # 첫 번째 탭 (데이터 테이블) 생성 및 추가
        self.tab1 = QWidget()
        self.tab_widget.addTab(self.tab1, "데이터 테이블")
        self.setup_table_tab()

        # 두 번째 탭 (변경 사항) 생성 및 추가
        self.tab2 = QWidget()
        self.tab_widget.addTab(self.tab2, "변경 사항")
        self.setup_result_tab()

    def setup_table_tab(self):
        """첫 번째 탭에 테이블과 확인 버튼을 설정합니다."""
        layout = QVBoxLayout(self.tab1)
        self.table_widget = QTableWidget()

        self.setup_table_content()

        # cellChanged 시그널을 연결하여 셀 내용 변경 시 포맷팅을 유지
        self.table_widget.cellChanged.connect(self.handle_cell_change)

        self.confirm_button = QPushButton("확인")
        self.confirm_button.clicked.connect(self.display_changes)

        layout.addWidget(self.table_widget)
        layout.addWidget(self.confirm_button)

    def setup_table_content(self):
        """테이블 위젯의 초기 내용을 설정합니다."""
        rows, cols = self.df.shape
        self.table_widget.setRowCount(rows)
        self.table_widget.setColumnCount(cols)

        self.table_widget.setHorizontalHeaderLabels(self.df.columns)

        for i in range(rows):
            for j in range(cols):
                header = self.df.columns[j]
                value = self.df.iloc[i, j]

                # '금액' 열은 천 단위 포맷팅 및 오른쪽 정렬 적용
                if header == '금액':
                    value_str = f'{value:,}'
                    item = QTableWidgetItem(value_str)
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                else:
                    item = QTableWidgetItem(str(value))

                self.table_widget.setItem(i, j, item)

    def setup_result_tab(self):
        """두 번째 탭에 변경 사항을 표시할 테이블을 설정합니다."""
        layout = QVBoxLayout(self.tab2)
        # 텍스트 에디터 대신 QTableWidget으로 변경
        self.result_table = QTableWidget()
        # 헤더 설정
        self.result_table.setColumnCount(4)
        self.result_table.setHorizontalHeaderLabels(['행', '열', '이전 값', '변경 값'])
        layout.addWidget(self.result_table)

    def handle_cell_change(self, row, column):
        """셀 내용이 변경될 때 호출되어 천 단위 포맷팅을 유지합니다."""
        if self.df.columns[column] == '금액':
            item = self.table_widget.item(row, column)
            if item is not None:
                new_value = item.text()

                # 콤마 제거 후 숫자로 변환 가능한지 확인
                clean_value = new_value.replace(',', '')
                try:
                    num_value = int(clean_value)
                    formatted_value = f'{num_value:,}'

                    # 무한 시그널 루프를 방지하기 위해 블록
                    self.table_widget.blockSignals(True)
                    new_item = QTableWidgetItem(formatted_value)
                    new_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    self.table_widget.setItem(row, column, new_item)
                    self.table_widget.blockSignals(False)

                except ValueError:
                    print(f"Warning: Invalid number format entered at row {row}, column {column}")

    def display_changes(self):
        """확인 버튼 클릭 시, 변경된 내용을 두 번째 탭의 테이블에 표시합니다."""
        changes = []
        rows = self.table_widget.rowCount()
        cols = self.table_widget.columnCount()

        for i in range(rows):
            for j in range(cols):
                header = self.df.columns[j]

                # 원본 DataFrame의 값을 사용하여 비교
                original_value = str(self.original_df.iloc[i, j])

                # 현재 테이블의 값을 가져오고 콤마를 제거하여 비교
                current_item = self.table_widget.item(i, j)
                current_value = current_item.text().replace(',', '') if current_item is not None else ""

                # 원본과 현재 값이 다를 경우 변경 사항에 추가
                if current_value != str(original_value):
                    changes.append({
                        'row': i + 1,
                        'column_name': header,
                        'original_value': original_value,
                        'new_value': current_item.text()
                    })

        # 결과 테이블 초기화
        self.result_table.setRowCount(0)

        if not changes:
            self.result_table.setRowCount(1)
            item = QTableWidgetItem("변경된 내용이 없습니다.")
            item.setTextAlignment(Qt.AlignCenter)
            self.result_table.setSpan(0, 0, 1, 4)
            self.result_table.setItem(0, 0, item)
            self.tab_widget.setCurrentIndex(1)
        else:
            self.result_table.setRowCount(len(changes))
            for i, change in enumerate(changes):
                # 변경 사항을 테이블의 각 셀에 삽입하고 빨간색으로 설정
                row_item = QTableWidgetItem(str(change['row']))
                col_item = QTableWidgetItem(change['column_name'])
                orig_item = QTableWidgetItem(str(change['original_value']))
                new_item = QTableWidgetItem(change['new_value'])

                # 텍스트 색상 설정 (setTextColor -> setForeground로 수정)
                new_item.setForeground(Qt.red)

                self.result_table.setItem(i, 0, row_item)
                self.result_table.setItem(i, 1, col_item)
                self.result_table.setItem(i, 2, orig_item)
                self.result_table.setItem(i, 3, new_item)

            # 두 번째 탭으로 전환
            self.tab_widget.setCurrentIndex(1)


def handle_exception(exc_type, exc_value, exc_traceback):
    """예외가 발생했을 때 호출되는 함수"""
    if issubclass(exc_type, KeyboardInterrupt):
        # Ctrl+C는 무시
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    # 예외 정보를 터미널에 출력
    print("\n\n--- Unhandled Exception ---")
    traceback.print_exception(exc_type, exc_value, exc_traceback)
    print("---------------------------\n")


if __name__ == '__main__':
    # 전역 예외 처리기를 등록
    sys.excepthook = handle_exception

    app = QApplication(sys.argv)

    # 예시 데이터프레임 생성
    data = {'상품명': ['사과', '바나나', '딸기'],
            '금액': [100000, 1500000, 5000]}
    df = pd.DataFrame(data)

    window = TableViewerWithTabs(df)
    window.show()
    sys.exit(app.exec_())
