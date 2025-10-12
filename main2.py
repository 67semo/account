import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QAction, QFileDialog, QTableWidget,
    QTableWidgetItem, QMessageBox, QVBoxLayout, QWidget, QMenu)
from PyQt5.QtCore import Qt
from subroutine_work import tax_calculation


class CsvEditor(QMainWindow):
    """
    CSV 파일 편집을 위한 GUI 애플리케이션 클래스입니다.
    """

    def __init__(self):
        super().__init__()
        self.setWindowTitle("CSV Editor")
        self.setGeometry(100, 100, 800, 600)

        self.table_widget = QTableWidget()
        self.current_file_path = None

        self.table_widget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table_widget.customContextMenuRequested.connect(self.show_context_menu)

        self.init_ui()

    def init_ui(self):
        """
        메인 UI를 초기화하고 메뉴바를 설정합니다.cd
        """
        # Central Widget 설정
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.addWidget(self.table_widget)

        # 메뉴바 설정
        menubar = self.menuBar()
        file_menu = menubar.addMenu("파일")

        # 파일 열기 (Open) 액션
        open_action = QAction("열기", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_file)
        file_menu.addAction(open_action)

        # 파일 저장 (Save) 액션
        save_action = QAction("저장", self)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.save_file)
        file_menu.addAction(save_action)

    def open_file(self):
        """
        파일 다이얼로그를 열어 CSV 파일을 선택하고 내용을 테이블에 로드합니다.
        """
        file_path, _ = QFileDialog.getOpenFileName(self, "Excel 파일 열기", "", "Excel Files (*.xls *.xlsx)")
        if file_path:
            self.current_file_path = file_path
            try:
                df, edit_df = tax_calculation.tax_invoice(file_path)
                self.load_table_from_dataframe(edit_df)
                self.setWindowTitle(f"경리 - {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "오류", f"파일을 읽는 중 오류가 발생했습니다: {e}")

    def load_table_from_dataframe(self, df):
        """
        pandas DataFrame의 내용을 QTableWidget에 표시합니다.
        """
        self.table_widget.setRowCount(df.shape[0])
        self.table_widget.setColumnCount(df.shape[1])
        self.table_widget.setHorizontalHeaderLabels(df.columns)

        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = QTableWidgetItem(str(df.iloc[i, j]))
                self.table_widget.setItem(i, j, item)

        self.table_widget.resizeColumnsToContents()

    def show_context_menu(self, pos):
        context_menu = QMenu(self)
        delete_action = QAction("행 삭제", self)
        delete_action.triggered.connect(self.delete_selected_rows)
        context_menu.addAction(delete_action)

        context_menu.exec_(self.table_widget.mapToGlobal(pos))

    def delete_selected_rows(self):
        selected_items = self.table_widget.selectedItems()

        if not selected_items:
            QMessageBox.warning(self, "경고", "삭제할 행을 선택해주세요.")
            return

        row_to_delete = {item.row() for item in selected_items}
        sorted_rows = sorted(list(row_to_delete), reverse=True)

        reply = QMessageBox.question(self, "확인",
                                     f"{len(sorted_rows)}개의 행을 삭제하시겠습니까?" ,
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            for row in sorted_rows:
                self.table_widget.removeRow(row)

    def save_file(self):
        """
        QTableWidget의 내용을 pandas DataFrame으로 변환하여 CSV 파일로 저장합니다.
        """
        if not self.current_file_path:
            QMessageBox.warning(self, "경고", "먼저 파일을 열어주세요.")
            return

        try:
            # QTableWidget의 내용을 DataFrame으로 변환
            rows = self.table_widget.rowCount()
            cols = self.table_widget.columnCount()
            headers = [self.table_widget.horizontalHeaderItem(j).text() for j in range(cols)]
            data = []
            for i in range(rows):
                row_data = [self.table_widget.item(i, j).text() if self.table_widget.item(i, j) else "" for j in
                            range(cols)]
                data.append(row_data)

            df = pd.DataFrame(data, columns=headers)

            # CSV 파일로 저장
            df.to_csv(self.current_file_path, index=False, encoding='utf-8-sig')
            QMessageBox.information(self, "성공", "파일이 성공적으로 저장되었습니다.")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"파일을 저장하는 중 오류가 발생했습니다: {e}")


def main():
    app = QApplication(sys.argv)
    editor = CsvEditor()
    editor.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
