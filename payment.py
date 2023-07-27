import pandas as pd
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QTableWidget, QTableWidgetItem, QLabel, \
    QPushButton, QHBoxLayout, QFileDialog, QMessageBox, QDialog
from qt_material import apply_stylesheet
import datetime
import logging
import os

def setup_logger():
    # 로그 파일 설정
    current_datetime = datetime.datetime.now()
    log_directory = '보상 지급 파일 제작 로그'
    log_filename = f'보상 지급 파일 제작 로그_{current_datetime.strftime("%Y%m%d_%H%M%S")}.txt'
    log_filepath = os.path.join(log_directory, log_filename)

    # logs 폴더가 없을 경우 생성
    if not os.path.exists(log_directory):
        os.makedirs(log_directory)

    # 기본 로거 초기화
    logging.basicConfig(level=logging.DEBUG)

    # 파일 핸들러 생성
    file_handler = logging.FileHandler(log_filepath)
    file_handler.setLevel(logging.DEBUG)

    # 포매터 생성
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    # 핸들러에 포매터 설정
    file_handler.setFormatter(formatter)

    # 로거에 핸들러 추가
    logger = logging.getLogger('')
    logger.addHandler(file_handler)
    return logger

# 기본 로거 초기화
logger = setup_logger()

class PopupDialog(QDialog):
    def __init__(self, parent=None, main_window=None):
        super().__init__(parent)
        self.setWindowTitle('Popup Dialog')
        self.main_window = main_window

        # 다이얼로그 레이아웃 설정
        layout = QVBoxLayout()
        self.setLayout(layout)

        # label 및 lineEdit 작성
        label_reward_type = QLabel('보상 종류 (Reward Type)')
        layout.addWidget(label_reward_type)
        self.reward_type_input = QtWidgets.QLineEdit()
        layout.addWidget(self.reward_type_input)

        label_reward_count = QLabel('보상 수량 (Reward Count)')
        layout.addWidget(label_reward_count)
        self.reward_count_input = QtWidgets.QLineEdit()
        layout.addWidget(self.reward_count_input)

        label_reward_info_id = QLabel('보상 정보 ID (Reward Info ID)')
        layout.addWidget(label_reward_info_id)
        self.reward_info_id_input = QtWidgets.QLineEdit()
        layout.addWidget(self.reward_info_id_input)

        label_item_bind = QLabel('귀속여부 (Item Bind)')
        layout.addWidget(label_item_bind)
        self.item_bind_input = QtWidgets.QLineEdit()
        self.item_bind_input.setText('1')  # Set default value
        layout.addWidget(self.item_bind_input)

        label_event_item_period_info_id = QLabel('아이템 사용 기간 Info ID (Event Item Period Info ID)')
        layout.addWidget(label_event_item_period_info_id)
        self.event_item_period_info_id_input = QtWidgets.QLineEdit()
        self.event_item_period_info_id_input.setText('0')  # Set default value
        layout.addWidget(self.event_item_period_info_id_input)

        # 버튼 생성
        self.add_reward_button = QPushButton('보상 추가', self)
        self.add_reward_button.clicked.connect(self.toMain)
        layout.addWidget(self.add_reward_button)

        self.clear_reward_button = QPushButton('종료', self)
        self.clear_reward_button.clicked.connect(self.Fending)
        layout.addWidget(self.clear_reward_button)

        # 스타일 입히기
        self.reward_type_input.setStyleSheet("QLineEdit { color: white; font-weight: bold; }")
        self.reward_count_input.setStyleSheet("QLineEdit { color: white; font-weight: bold; }")
        self.reward_info_id_input.setStyleSheet("QLineEdit { color: white; font-weight: bold; }")
        self.item_bind_input.setStyleSheet("QLineEdit { color: white; font-weight: bold; }")
        self.event_item_period_info_id_input.setStyleSheet("QLineEdit { color: white; font-weight: bold; }")
        label_reward_type.setStyleSheet("QLabel { font-weight: bold; }")
        label_reward_count.setStyleSheet("QLabel { font-weight: bold; }")
        label_reward_info_id.setStyleSheet("QLabel { font-weight: bold; }")
        label_item_bind.setStyleSheet("QLabel { font-weight: bold; }")
        label_event_item_period_info_id.setStyleSheet("QLabel { font-weight: bold; }")

    def toMain(self):
        # 메인 윈도우에 값 전달
        if (self.reward_type_input.text().strip() == "" or
                self.reward_count_input.text().strip() == "" or
                self.item_bind_input.text().strip() == "" or
                self.event_item_period_info_id_input.text().strip() == ""):
            QMessageBox.warning(self, '경고', '보상 값을 확인해 주세요.\n'
                                            '보상 종류, 수량, 귀속여부, 아이템 사용 기간은 필수 조건입니다.')
        else:
            self.main_window.add_reward_to_table(self.reward_type_input.text(), self.reward_count_input.text(), self.reward_info_id_input.text(), self.item_bind_input.text(), self.event_item_period_info_id_input.text())
    def Fending(self):
        # 다이얼 로그 종료
        self.accept()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('보상 지급 엑셀파일 제작')
        self.setGeometry(0,0,600,800)
        self.layout = QVBoxLayout()
        self.central_widget = QWidget()
        self.central_widget.setLayout(self.layout)
        self.setCentralWidget(self.central_widget)

        # 메인 윈도우 UI
        label_server_receiver = QLabel('서버ID & CID 예)1301:6280560334140862464,')
        self.layout.addWidget(label_server_receiver)

        input_layout = QHBoxLayout()
        self.layout.addLayout(input_layout)

        # textEdit
        self.server_receiver_input = QtWidgets.QTextEdit()
        input_layout.addWidget(self.server_receiver_input)

        # tableWidget1
        label_reward_table = QLabel('보상 정보')
        self.layout.addWidget(label_reward_table)
        self.reward_table = QTableWidget()
        self.reward_table.setColumnCount(5)
        self.reward_table.setHorizontalHeaderLabels(['보상 종류', '보상 수량', '보상 정보 ID', '귀속여부', '아이템 사용 기간 Info ID'])
        self.layout.addWidget(self.reward_table)

        input_box = QtWidgets.QHBoxLayout()
        self.layout.addLayout(input_box)

        # pushButton1, 2
        self.dialog_button = QPushButton('보상 추가', self)
        self.dialog_button.clicked.connect(self.dialog_exec)
        input_box.addWidget(self.dialog_button)
        self.clear_reward_button = QPushButton('보상 초기화', self)
        self.clear_reward_button.clicked.connect(self.clear_reward_table)
        input_box.addWidget(self.clear_reward_button)

        # tableWidget2
        label_result_table = QLabel('미리 보기')
        self.layout.addWidget(label_result_table)
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(7)
        self.result_table.setHorizontalHeaderLabels(['서버 ID', '수신인 ID', '보상 종류',
                                                    '보상 수량', '보상 정보 ID',
                                                    '귀속여부', '아이템 사용 기간 ID'])
        self.layout.addWidget(self.result_table)

        input_layout2 = QHBoxLayout()
        self.layout.addLayout(input_layout2)

        # pushButton3, 4
        self.add_table_button = QPushButton('테이블 결합', self)
        self.add_table_button.clicked.connect(self.add_table_data)
        input_layout2.addWidget(self.add_table_button)
        self.clear_button = QPushButton('테이블 초기화', self)
        self.clear_button.clicked.connect(self.clear_table)
        input_layout2.addWidget(self.clear_button)

        input_layout3 = QHBoxLayout()
        self.layout.addLayout(input_layout3)

        # pushButton5, 6
        self.export_button = QPushButton('엑셀 파일 추출', self)
        self.export_button.clicked.connect(self.export_to_excel)
        input_layout3.addWidget(self.export_button)
        self.close_button = QPushButton('종료', self)
        self.close_button.clicked.connect(self.close_app)
        input_layout3.addWidget(self.close_button)
        self.setup_ui()

        # 스타일 입히기
        label_server_receiver.setStyleSheet("QLabel { font-weight: bold; }")
        self.reward_table.setStyleSheet("QTableWidget { color: white; font-weight: bold; }")
        label_reward_table.setStyleSheet("QLabel { font-weight: bold; }")
        self.result_table.setStyleSheet("QTableWidget { color: white; font-weight: bold; }")
        label_result_table.setStyleSheet("QLabel { font-weight: bold; }")

        # textedit 관련 함수
        self.server_receiver_input.textChanged.connect(self.duplicate)
        self.server_receiver_input.setFixedHeight(100)

        # 데이터 프레임 뼈대 초기화
        self.existing_dataframe = pd.DataFrame(columns=['서버 ID', '수신인 ID', '보상 종류', '보상 수량',
                                                        '보상 정보 ID', '귀속여부', '아이템 사용 기간 Info ID'])

        # 테이블 정렬
        self.result_table.resizeColumnsToContents()
        self.result_table.resizeRowsToContents()
        self.reward_table.resizeColumnsToContents()
        self.reward_table.resizeRowsToContents()

    def setup_ui(self):
        self.show()

    def duplicate(self):
        # 중복 CID 검사 함수
        input_text = self.server_receiver_input.toPlainText()
        lines = input_text.split('\n')
        has_duplicates = False
        has_duplicates_string = []
        for i in range(len(lines)):
            for j in range(i + 1, len(lines)):
                if lines[i] != '' and lines[i] == lines[j]:
                    has_duplicates = True
                    has_duplicates_string.append(lines[i])
                    has_duplicates_string.append(lines[j])
                    break
            if has_duplicates:
                break
        if has_duplicates:
            QMessageBox.warning(self, '경고', f'CID 중복이 확인 되었습니다.{has_duplicates_string}')
            logger.warning(f'CID 중복이 확인 되었습니다.{has_duplicates_string}')

    def add_reward_to_table(self, a, b, c, d, e):
        # 다이얼로그 값 받아와 테이블에 추가
        row_count = self.reward_table.rowCount()
        self.reward_table.insertRow(row_count)
        self.reward_table.setItem(row_count, 0, QTableWidgetItem(a))
        self.reward_table.setItem(row_count, 1, QTableWidgetItem(b))
        self.reward_table.setItem(row_count, 2, QTableWidgetItem(c))
        self.reward_table.setItem(row_count, 3, QTableWidgetItem(d))
        self.reward_table.setItem(row_count, 4, QTableWidgetItem(e))

        self.reward_table.resizeColumnsToContents()
        self.reward_table.resizeRowsToContents()

        logger.info(f'보상 추가 완료. 보상 종류: {a}, 보상 수량: {b}, '
                    f'보상 정보 ID: {c}, 귀속여부: {d}, '
                    f'아이템 사용 기간 Info ID: {e}')

    def clear_reward_table(self):
        # 보상 테이블 초기화
        if self.reward_table.rowCount() > 0:
            self.reward_table.clearContents()
            self.reward_table.setRowCount(0)
            logger.info('보상 테이블 초기화 완료.')

    def add_table_data(self):
        # 서버ID, CID, 보상 테이블 결합
        if self.server_receiver_input.toPlainText().strip() != "" and self.reward_table.rowCount() > 0:
            server_receiver_text = self.server_receiver_input.toPlainText()
            server_receiver_lines = server_receiver_text.split('\n')

            reward_data = []
            # reward_table 값 리스트 화
            for row in range(self.reward_table.rowCount()):
                reward_type = self.reward_table.item(row, 0).text()
                reward_count = self.reward_table.item(row, 1).text()
                reward_info_id = self.reward_table.item(row, 2).text()
                item_bind = self.reward_table.item(row, 3).text()
                event_item_period_info_id = self.reward_table.item(row, 4).text()
                reward_data.append([reward_type, reward_count, reward_info_id, item_bind, event_item_period_info_id])

            data = []
            # server_receiver_input 값 리스트 화 및 결합
            for line in server_receiver_lines:
                line = line.strip()
                if line:
                    parts = line.split(':')
                    if len(parts) == 2:
                        server_id = parts[0].strip()
                        receiver_id = parts[1].strip(',')

                        for reward in reward_data:
                            data.append([server_id, receiver_id] + reward)

            # result_table 생성 및 값 입력
            row_count = self.result_table.rowCount()
            self.result_table.setRowCount(row_count + len(data))
            for row, row_data in enumerate(data):
                for col, value in enumerate(row_data):
                    item = QTableWidgetItem(str(value))
                    self.result_table.setItem(row_count + row, col, item)

            self.result_table.resizeColumnsToContents()
            self.result_table.resizeRowsToContents()
            logger.info(f'테이블 추가 완료. 행 개수 : {len(data)}')

    def clear_table(self):
        # 결합 테이블 초기화
        if self.result_table.rowCount() > 0:
            self.result_table.clearContents()
            self.result_table.setRowCount(0)
            logger.info('미리 보기 테이블 초기화 완료.')

    def export_to_excel(self):
        # 엑셀 파일 추출
        if self.result_table.rowCount() == 0:
            QMessageBox.warning(self, '경고', '테이블에 데이터가 없습니다.')
            return

        file_dialog = QFileDialog(self)
        file_dialog.setAcceptMode(QFileDialog.AcceptSave)
        file_dialog.setNameFilter('Excel Files (*.xlsx)')
        file_dialog.setDefaultSuffix('xlsx')
        if file_dialog.exec() != QFileDialog.Accepted:
            return

        file_path = file_dialog.selectedFiles()[0]
        if not file_path:
            return

        # 예외 처리
        try:
            df = self.get_table_data()
            writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='Sheet1')

            workbook = writer.book
            worksheet = workbook.get_worksheet_by_name('Sheet1')
            text_format = workbook.add_format({'num_format': '@'})

            for col in range(df.shape[1]):
                worksheet.write(0, col, str(df.columns[col]), text_format)

            for row in range(1, df.shape[0] + 1):
                for col in range(df.shape[1]):
                    value = df.iloc[row - 1, col]
                    if pd.isnull(value):
                        worksheet.write(row, col, '', text_format)
                    else:
                        worksheet.write(row, col, str(value), text_format)
            writer.close()

            QMessageBox.information(self, '알림', '엑셀 파일이 저장되었습니다.')
            logger.info(f'엑셀 파일 저장 완료. 저장 경로: {file_path}')
        except Exception as e:
            QMessageBox.warning(self, '오류', f'엑셀 파일 저장 중 오류가 발생했습니다: {str(e)}')
            logger.warning(f'엑셀 파일 저장 중 오류 발생. 오류 내용: {str(e)}, {e}')

    def get_table_data(self):
        # 테이블 데이터프레임 화
        rows = self.result_table.rowCount()
        rows = rows + 5

        # 기존 데이터프레임 뼈대 가져오기
        data = {
            'server_id': [],
            'receiver': [],
            'reward_type': [],
            'reward_count': [],
            'reward_info_id': [],
            'item_bind': [],
            'event_item_period_info_id': []
        }

        df = pd.DataFrame(data)

        # 기존 데이터프레임 뼈대의 2, 3, 4행 유지
        df.loc[2] = ['', '', '', '', '', '', '']
        df.loc[3] = [
            '서버 ID', '수신인 ID (Player OR User)', '보상 종류', '보상 수량',
            '보상 정보 ID', '귀속여부', '아이템 사용 기간 Info ID'
        ]
        df.loc[4, ['reward_type', 'reward_count']] = '0'

        # 테이블 위젯의 값 추가
        for row in range(5, rows):
            row_data = []
            for col in range(self.result_table.columnCount()):
                item = self.result_table.item(row - 5, col)
                if item is not None:
                    row_data.append(item.text())
                else:
                    row_data.append('')
            df.loc[row] = row_data
            logger.info(f'행 삽입 완료. {row}행 입력 데이터 : {row_data}')

        # 기존 데이터프레임 뼈대와 새로 추가한 데이터프레임 결합
        self.existing_dataframe = pd.concat([self.existing_dataframe, df], ignore_index=True)

        logger.info('데이터 프레임 생성 및 전달 완료')
        return df

    def close_app(self):
        # 앱 종료
        self.close()
        logger.info('앱 종료')

    def dialog_exec(self):
        # 다이얼로그 실행
        dlg = PopupDialog(self, main_window=self)
        dlg.exec_()

# Create PyQt5 application
app = QApplication([])
apply_stylesheet(app, theme='dark_teal.xml')
window = MainWindow()
app.exec()