import pandas as pd
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QTableWidget, QTableWidgetItem, QLabel, \
    QPushButton, QHBoxLayout, QFileDialog, QMessageBox
from qt_material import apply_stylesheet
import datetime
import logging
import os


def setup_logger():
    # 로그 파일 설정
    current_datetime = datetime.datetime.now()
    log_directory = '내부재화지급 파일 제작 로그'
    log_filename = f'내부재화지급 파일 제작 로그_{current_datetime.strftime("%Y%m%d_%H%M%S")}.txt'
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

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('내부재화 지급 파일 제작 앱')
        self.setGeometry(0, 0, 600, 800)
        self.layout = QVBoxLayout()
        self.central_widget = QWidget()
        self.central_widget.setLayout(self.layout)
        self.setCentralWidget(self.central_widget)

        # UI
        label_server_receiver = QLabel('서버ID & CID 예)1301:6280560334140862464,')
        self.layout.addWidget(label_server_receiver)

        input_layout = QHBoxLayout()
        self.layout.addLayout(input_layout)

        self.server_receiver_input = QtWidgets.QTextEdit()
        input_layout.addWidget(self.server_receiver_input)

        input_box = QtWidgets.QVBoxLayout()
        input_layout.addLayout(input_box)

        self.add_reward_button_1 = QPushButton('무과금', self)
        self.add_reward_button_1.clicked.connect(self.add_reward_row)
        input_box.addWidget(self.add_reward_button_1)

        self.add_reward_button_2 = QPushButton('저과금', self)
        self.add_reward_button_2.clicked.connect(self.add_reward_middle)
        input_box.addWidget(self.add_reward_button_2)

        self.add_reward_button_3 = QPushButton('고과금', self)
        self.add_reward_button_3.clicked.connect(self.add_reward_high)
        input_box.addWidget(self.add_reward_button_3)

        self.add_reward_button_4 = QPushButton('초고과금', self)
        self.add_reward_button_4.clicked.connect(self.add_reward_veryhigh)
        input_box.addWidget(self.add_reward_button_4)

        self.add_reward_button_5 = QPushButton('고과금(주차 보상)', self)
        self.add_reward_button_5.clicked.connect(self.add_reward_high_weekly)
        input_box.addWidget(self.add_reward_button_5)

        self.add_reward_button_6 = QPushButton('초고과금(주차 보상)', self)
        self.add_reward_button_6.clicked.connect(self.add_reward_veryhigh_weekly)
        input_box.addWidget(self.add_reward_button_6)

        self.clear_button = QPushButton('테이블 초기화', self)
        self.clear_button.clicked.connect(self.clear_table)
        input_box.addWidget(self.clear_button)

        label_result_table = QLabel('미리 보기')
        self.layout.addWidget(label_result_table)
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(7)
        self.result_table.setHorizontalHeaderLabels(['서버 ID', '수신인 ID', '보상 종류',
                                                     '보상 수량', '보상 정보 ID',
                                                     '귀속여부', '아이템 사용 기간 ID'])
        self.layout.addWidget(self.result_table)

        self.export_button = QPushButton('엑셀 파일 추출', self)
        self.export_button.clicked.connect(self.export_to_excel)
        self.layout.addWidget(self.export_button)

        self.close_button = QPushButton('종료', self)
        self.close_button.clicked.connect(self.close_app)
        self.layout.addWidget(self.close_button)

        self.setup_ui()

        # 스타일 입히기
        label_server_receiver.setStyleSheet("QLabel { font-weight: bold; }")
        self.result_table.setStyleSheet("QTableWidget { color: white; font-weight: bold; }")
        label_result_table.setStyleSheet("QLabel { font-weight: bold; }")

        # textedit 중복 검사 시그널
        self.server_receiver_input.textChanged.connect(self.duplicate)

        # 데이터 프레임 뼈대
        self.existing_dataframe = pd.DataFrame(columns=['서버 ID', '수신인 ID', '보상 종류', '보상 수량',
                                                        '보상 정보 ID', '귀속여부', '아이템 사용 기간 Info ID'])

        # 테이블 정렬
        self.result_table.resizeColumnsToContents()

        # 텍스트 에딧 크기 조정
        self.server_receiver_input.setFixedHeight(300)

    def setup_ui(self):
        self.show()

    def duplicate(self):
        # CID 중복 검사 함수
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
            logger.warning(f'CID 중복 확인 {has_duplicates_string}')

    def add_reward_row(self):
        # 무과금 보상
        coupon = ['6', '300', '10073', '1', '0']
        reward = [coupon]
        self.add_result_row(reward)

    def add_reward_middle(self):
        # 저과금 보상
        coupon = ['6', '2200', '10073', '1', '0']
        dailyDia = ['6', '1', '930190', '1', '0']
        dailyQuest = ['6', '1', '930203', '1', '0']
        dailyHunt = ['6', '1', '930204', '1', '0']
        dailyItem = ['6', '1', '930205', '1', '0']
        revolPlus = ['6', '1', '110910182', '1', '0']
        freeDia = ['2', '5000', '0', '1', '0']
        reward = [coupon, dailyDia, dailyQuest, dailyHunt, dailyItem, revolPlus]

        for i in range(4):
            reward.append(freeDia[:])
        self.add_result_row(reward)

    def add_reward_high(self):
        # 고과금 보상
        coupon = ['6', '1100', '10073', '1', '0']
        dailyDia = ['6', '1', '930190', '1', '0']
        dailyQuest = ['6', '1', '930203', '1', '0']
        dailyHunt = ['6', '1', '930204', '1', '0']
        dailyItem = ['6', '1', '930205', '1', '0']
        revolPlus = ['6', '1', '110910182', '1', '0']
        dailyHarvest = ['6', '1', '930193', '1', '0']
        dailySilen = ['6', '1', '931547', '1', '0']
        adenaBox = ['6', '10', '920206', '1', '0']
        freeDia = ['2', '5000', '0', '1', '0']
        reward = [coupon, dailyDia, dailyQuest, dailyHunt, dailyItem, revolPlus, dailyHarvest, dailySilen, adenaBox]

        for i in range(10):
            reward.append(freeDia[:])
        self.add_result_row(reward)

    def add_reward_veryhigh(self):
        # 초고과금 보상
        coupon = ['6', '1500', '10073', '1', '0']
        dailyDia = ['6', '1', '930190', '1', '0']
        dailyQuest = ['6', '1', '930203', '1', '0']
        dailyHunt = ['6', '1', '930204', '1', '0']
        dailyItem = ['6', '1', '930205', '1', '0']
        revolPlus = ['6', '1', '110910182', '1', '0']
        dailyHarvest = ['6', '1', '930193', '1', '0']
        dailySilen = ['6', '1', '931547', '1', '0']
        adenaBox = ['6', '10', '920206', '1', '0']
        freeDia = ['2', '5000', '0', '1', '0']
        reward = [coupon, dailyDia, dailyQuest, dailyHunt, dailyItem, revolPlus, dailyHarvest, dailySilen, adenaBox]

        for i in range(10):
            reward.append(freeDia[:])
        self.add_result_row(reward)

    def add_reward_high_weekly(self):
        # 고과금 주차별 보상
        coupon = ['6', '1100', '10073', '1', '0']
        adenaBox = ['6', '10', '920206', '1', '0']
        freeDia = ['2', '5000', '0', '1', '0']
        reward = [coupon, adenaBox]

        for i in range(10):
            reward.append(freeDia[:])
        self.add_result_row(reward)

    def add_reward_veryhigh_weekly(self):
        # 초고과금 주차별 보상
        coupon = ['6', '1500', '10073', '1', '0']
        adenaBox = ['6', '10', '920206', '1', '0']
        freeDia = ['2', '5000', '0', '1', '0']
        reward = [coupon, adenaBox]

        for i in range(10):
            reward.append(freeDia[:])
        self.add_result_row(reward)

    def add_result_row(self, reward):
        # 테이블에 미리 보기 추가
        reask_server_id = 0
        reask_receiver_id = 0
        if self.server_receiver_input.toPlainText().strip() != "":
            self.result_table.setRowCount(0)
            server_receiver_text = self.server_receiver_input.toPlainText()
            server_receiver_lines = server_receiver_text.split('\n')

            data = []
            for line in server_receiver_lines:
                line = line.strip()
                if line:
                    parts = line.split(':')
                    if len(parts) == 2:
                        server_id = parts[0].strip()
                        receiver_id = parts[1].strip(',')
                        if len(server_id) != 4 and reask_server_id == 0:
                            result = QMessageBox.question(self, "서버ID 확인 필요", f"{server_id}의 서버ID가 4자릿수가 아닙니다. 계속 진행 하시겠습니까?\n"
                                                                              f"예 버튼 클릭 시 두번 다시 묻지 않음.",
                                                          QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                            logger.info(f'예외 서버ID 확인. 서버ID : {server_id}')
                            if result == QMessageBox.Yes:
                                reask_server_id += 1
                                logger.info('예외 서버ID 허용.')
                            if result == QMessageBox.No:
                                logger.info('보상 선택 취소.')
                                return
                        if len(receiver_id) != 19 and reask_receiver_id == 0:
                            result2 = QMessageBox.question(self, "CID(UID) 확인 필요",
                                                           f"{receiver_id}의 CID(UID)가 19자릿수가 아닙니다. 계속 진행 하시겠습니까?\n"
                                                           f"예 버튼 클릭 시 두번 다시 묻지 않음.",
                                                           QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                            logger.info(f'예외 CID(UID) 확인. CID(UID) : {receiver_id}')
                            if result2 == QMessageBox.Yes:
                                reask_receiver_id += 1
                                logger.info('예외 CID(UID) 허용.')
                            if result2 == QMessageBox.No:
                                logger.info('보상 선택 취소.')
                                return
                        for i in reward:
                            data.append([server_id, receiver_id] + i)

            for i in range(len(data)):
                row_count = self.result_table.rowCount()
                self.result_table.insertRow(row_count)
                for col, reward_data in enumerate(data[i]):
                    item = QTableWidgetItem(reward_data)
                    self.result_table.setItem(row_count, col, item)
        if reward[0][1] == '300':
            logger.info('무과금 보상 선택')
        elif reward[0][1] == '2200':
            logger.info('저과금 보상 선택')
        elif reward[0][1] == '1100' and self.result_table.rowCount() == 12:
            logger.info('고과금 주차별 보상 선택')
        elif reward[0][1] == '1500' and self.result_table.rowCount() == 12:
            logger.info('초고과금 주차별 보상 선택')
        elif reward[0][1] == '1100':
            logger.info('고과금 보상 선택')
        elif reward[0][1] == '1500':
            logger.info('초고과금 보상 선택')

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
            self.result_table.setRowCount(0)
            logger.info(f'엑셀 파일 저장 완료. 저장 경로: {file_path}')
        except Exception as e:
            QMessageBox.warning(self, '오류', f'엑셀 파일 저장 중 오류가 발생했습니다: {str(e)}')
            logger.warning(f'엑셀 파일 저장 중 오류 발생. 오류 내용: {str(e)}, {e}')

    def get_table_data(self):
        # 테이블 데이터프레임화
        rows = self.result_table.rowCount()
        rows = rows + 5

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
        header = [
            '서버 ID', '수신인 ID (Player OR User)', '보상 종류', '보상 수량',
            '보상 정보 ID', '귀속여부', '아이템 사용 기간 Info ID'
        ]
        df.loc[2] = ['', '', '', '', '', '', '']
        df.loc[3] = [
            '서버 ID', '수신인 ID (Player OR User)', '보상 종류', '보상 수량',
            '보상 정보 ID', '귀속여부', '아이템 사용 기간 Info ID'
        ]
        df.loc[4, ['reward_type', 'reward_count']] = '0'

        logger.info(f'기존 4행 추가 완료. 테이블내 데이터 삽입 시작. 헤더 데이터 : {header}')
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

        logger.info(f'전체 삽입 행 : {self.result_table.rowCount()}, 데이터 프레임 생성 및 전달 완료')
        return df

    def close_app(self):
        # 앱 종료
        self.close()
        logger.info('앱 종료')


# Create PyQt5 application
app = QApplication([])
apply_stylesheet(app, theme='dark_teal.xml')
window = MainWindow()
app.exec()
