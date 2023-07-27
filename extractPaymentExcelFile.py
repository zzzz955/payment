import sys
import json
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, QTabWidget, QTextEdit, \
    QLabel, QHBoxLayout, QFileDialog, QMessageBox, QDialog, QTableWidget, QTableWidgetItem, QInputDialog
from PyQt5 import QtWidgets
from qt_material import apply_stylesheet
from log import setup_logger
import pandas as pd

logger = setup_logger()

class PopupDialog(QDialog):
    # 보상 추가 다이얼로그
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('보상 추가')
        self.main_tab = parent

        layout = QVBoxLayout(self)

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
        self.item_bind_input.setText('1')  # 기본 값 설정
        layout.addWidget(self.item_bind_input)

        label_event_item_period_info_id = QLabel('아이템 사용 기간 Info ID (Event Item Period Info ID)')
        layout.addWidget(label_event_item_period_info_id)
        self.event_item_period_info_id_input = QtWidgets.QLineEdit()
        self.event_item_period_info_id_input.setText('0')  # 기본 값 설정
        layout.addWidget(self.event_item_period_info_id_input)

        self.add_reward_button = QPushButton('보상 추가', self)
        self.add_reward_button.clicked.connect(self.add_reward_to_table_dialog)
        layout.addWidget(self.add_reward_button)

        self.clear_reward_button = QPushButton('종료', self)
        self.clear_reward_button.clicked.connect(self.close)
        layout.addWidget(self.clear_reward_button)

        # 다이얼로그 위젯 스타일 입히기
        label_reward_type.setStyleSheet("QLabel { font-weight: bold; }")
        label_reward_count.setStyleSheet("QLabel { font-weight: bold; }")
        label_reward_info_id.setStyleSheet("QLabel { font-weight: bold; }")
        label_item_bind.setStyleSheet("QLabel { font-weight: bold; }")
        label_event_item_period_info_id.setStyleSheet("QLabel { font-weight: bold; }")
        self.reward_type_input.setStyleSheet("QLineEdit { color: white; font-weight: bold; }")
        self.reward_count_input.setStyleSheet("QLineEdit { color: white; font-weight: bold; }")
        self.reward_info_id_input.setStyleSheet("QLineEdit { color: white; font-weight: bold; }")
        self.item_bind_input.setStyleSheet("QLineEdit { color: white; font-weight: bold; }")
        self.event_item_period_info_id_input.setStyleSheet("QLineEdit { color: white; font-weight: bold; }")

    def add_reward_to_table_dialog(self):
        # 메인탭 보상 테이블에 정보 추가하는 함수
        reward_type = self.reward_type_input.text()
        reward_count = self.reward_count_input.text()
        reward_info_id = self.reward_info_id_input.text()
        item_bind = self.item_bind_input.text()
        event_item_period_info_id = self.event_item_period_info_id_input.text()

        # 필수 값 입력 여부 체크
        if not (reward_type.strip() and reward_count.strip() and item_bind.strip() and event_item_period_info_id.strip()):
            QMessageBox.warning(self, '경고', '보상 값을 확인해 주세요.\n'
                                            '보상 종류, 수량, 귀속여부, 아이템 사용 기간은 필수 조건입니다.')
            return

        # 메인탭 테이블에 값 추가 함수 호출
        self.main_tab.add_reward_to_table(reward_type, reward_count, reward_info_id, item_bind, event_item_period_info_id)

class MainTab(QWidget):
    # 메인 탭
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.layout = QVBoxLayout(self)

        label_server_receiver = QLabel('지급 대상자 정보 예)1301:6280560334140862464,')
        self.layout.addWidget(label_server_receiver)

        input_layout = QHBoxLayout()
        self.layout.addLayout(input_layout)

        # 유저 정보 입력 텍스트 에딧
        self.server_receiver_input = QtWidgets.QTextEdit()
        input_layout.addWidget(self.server_receiver_input)
        self.server_receiver_input.textChanged.connect(self.duplicate)
        self.server_receiver_input.setFixedHeight(100)

        # 보상 정보 테이블
        label_reward_table = QLabel('보상 정보')
        self.layout.addWidget(label_reward_table)
        self.reward_table = QTableWidget()
        self.reward_table.setColumnCount(6)
        self.reward_table.setHorizontalHeaderLabels(['삭제', '보상 종류', '보상 수량', '보상 정보 ID', '귀속여부', '아이템 사용 기간 Info ID'])
        self.layout.addWidget(self.reward_table)
        self.reward_table.resizeColumnsToContents()
        self.reward_table.resizeRowsToContents()

        input_box = QHBoxLayout()
        self.layout.addLayout(input_box)

        # 보상 관련 버튼
        self.dialog_button = QPushButton('보상 추가', self)
        self.dialog_button.clicked.connect(self.dialog_exec)
        input_box.addWidget(self.dialog_button)
        self.clear_reward_button = QPushButton('보상 초기화', self)
        self.clear_reward_button.clicked.connect(self.clear_reward_table)
        input_box.addWidget(self.clear_reward_button)

        # 엑셀 추출 버튼
        self.export_button = QPushButton('엑셀 파일 추출', self)
        self.export_button.clicked.connect(self.export_to_excel)
        input_box.addWidget(self.export_button)

        # 메인탭 위젯 스타일 입히기
        label_server_receiver.setStyleSheet("QLabel { font-weight: bold; }")
        self.reward_table.setStyleSheet("QTableWidget { color: white; font-weight: bold; }")
        label_reward_table.setStyleSheet("QLabel { font-weight: bold; }")

        # 데이터 프레임 헤더 설정
        self.existing_dataframe = pd.DataFrame(columns=['서버 ID', '수신인 ID', '보상 종류', '보상 수량',
                                                        '보상 정보 ID', '귀속여부', '아이템 사용 기간 Info ID'])
        # 가상 테이블 (유저 및 보상 정보 결합용)
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(7)
        self.result_table.setHorizontalHeaderLabels(['서버 ID', '수신인 ID', '보상 종류',
                                                     '보상 수량', '보상 정보 ID',
                                                     '귀속여부', '아이템 사용 기간 ID'])


    def dialog_exec(self):
        # 보상 추가 다이얼로그 실행 함수
        dlg = PopupDialog(self)
        dlg.exec_()

    def add_reward_to_table(self, reward_type, reward_count, reward_info_id, item_bind, event_item_period_info_id):
        # 보상 추가 다이얼로그 데이터를 받아 테이블에 정보를 추가 하는 함수
        row_count = self.reward_table.rowCount()
        self.reward_table.insertRow(row_count)
        delete_button = QPushButton('삭제')
        delete_button.clicked.connect(self.delete_row)
        self.reward_table.setCellWidget(row_count, 0, delete_button)
        self.reward_table.setItem(row_count, 1, QTableWidgetItem(reward_type))
        self.reward_table.setItem(row_count, 2, QTableWidgetItem(reward_count))
        self.reward_table.setItem(row_count, 3, QTableWidgetItem(reward_info_id))
        self.reward_table.setItem(row_count, 4, QTableWidgetItem(item_bind))
        self.reward_table.setItem(row_count, 5, QTableWidgetItem(event_item_period_info_id))
        self.reward_table.resizeColumnsToContents()
        self.reward_table.resizeRowsToContents()

    def delete_row(self):
        # 행 삭제 함수
        button = self.sender()
        if button:
            index = self.reward_table.indexAt(button.pos())
            if index.isValid():
                row = index.row()
                self.reward_table.removeRow(row)

    def clear_reward_table(self):
        # 보상 테이블 초기화 함수
        if self.reward_table.rowCount() > 0:
            self.reward_table.clearContents()
            self.reward_table.setRowCount(0)
            logger.info('보상 테이블 초기화 완료.')

    def duplicate(self):
        # 유저 정보 중복 검사 함수
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
            QMessageBox.warning(self, '경고', f'유저 정보 중복이 확인 되었습니다.{has_duplicates_string}')
            logger.warning(f'유저 정보 중복이 확인 되었습니다.{has_duplicates_string}')

    def export_to_excel(self):
        # 엑셀 파일 추출 함수
        # 보상 및 유저 정보가 공란일 경우 리턴
        if self.reward_table.rowCount() == 0 or len(self.server_receiver_input.toPlainText()) == 0:
            QMessageBox.warning(self, '경고', '테이블에 데이터가 없습니다.')
            return

        # 파일 다이얼로그 실행
        file_dialog = QFileDialog(self)
        file_dialog.setAcceptMode(QFileDialog.AcceptSave)
        file_dialog.setNameFilter('Excel Files (*.xlsx)')
        file_dialog.setDefaultSuffix('xlsx')
        if file_dialog.exec() != QFileDialog.Accepted:
            return

        # 파일 경로 부재시 리턴
        file_path = file_dialog.selectedFiles()[0]
        if not file_path:
            return

        # 예외 처리
        try:
            # 데이터프레임 엑셀화
            df = self.get_table_data()
            if df is None:
                QMessageBox.information(self, '정보', f'엑셀 파일 저장을 중단 하였습니다.')
                return
            writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='Sheet1')

            # 엑셀 데이터 텍스트 형식으로 변경
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
        # 가상 테이블 생성 및 데이터프레임 화 함수
        reask_server_id = 0
        reask_receiver_id = 0

        # 유저 정보 가져오기
        if self.server_receiver_input.toPlainText().strip() != "":
            self.result_table.setRowCount(0)
            server_receiver_text = self.server_receiver_input.toPlainText()
            server_receiver_lines = server_receiver_text.split('\n')

            # 보상 정보 가져오기
            reward_data = []
            for row in range(self.reward_table.rowCount()):
                reward_type = self.reward_table.item(row, 1).text()
                reward_count = self.reward_table.item(row, 2).text()
                reward_info_id = self.reward_table.item(row, 3).text()
                item_bind = self.reward_table.item(row, 4).text()
                event_item_period_info_id = self.reward_table.item(row, 5).text()
                reward_data.append([reward_type, reward_count, reward_info_id, item_bind, event_item_period_info_id])

            # 유저 정보 특이사항 체크
            editdata = []
            for line in server_receiver_lines:
                line = line.strip()
                if line:
                    parts = line.split(':')
                    if len(parts) == 2:
                        server_id = parts[0].strip()
                        receiver_id = parts[1].strip(',')

                        # 서버ID 4자리 여부 확인
                        if len(server_id) != 4 and reask_server_id == 0:
                            result = QMessageBox.question(self, "서버ID 확인 필요",
                                                          f"{server_id}의 서버ID가 4자릿수가 아닙니다. 계속 진행 하시겠습니까?\n"
                                                          f"예 버튼 클릭 시 두번 다시 묻지 않음.",
                                                          QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                            logger.info(f'예외 서버ID 확인. 서버ID : {server_id}')
                            if result == QMessageBox.Yes:
                                reask_server_id += 1
                                logger.info('예외 서버ID 허용.')
                            if result == QMessageBox.No:
                                logger.info('예외 서버ID 허용 불가, 보상 선택 취소.')
                                return

                        # CID 19자리 여부 확인
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
                                logger.info('예외 CID(UID) 허용 불가, 보상 선택 취소.')
                                return
                        # 유저 및 보상 정보 리스트 결합
                        for reward in reward_data:
                            editdata.append([server_id, receiver_id]+reward)

            row_count = self.result_table.rowCount()
            self.result_table.setRowCount(row_count + len(editdata))
            for row, row_data in enumerate(editdata):
                for col, value in enumerate(row_data):
                    item = QTableWidgetItem(str(value))
                    self.result_table.setItem(row_count + row, col, item)

            self.result_table.resizeColumnsToContents()
            self.result_table.resizeRowsToContents()
            logger.info(f'테이블 추가 완료. 행 개수 : {len(editdata)}')

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

    def to_json(self):
        # reward_table의 보상 정보를 JSON 형식으로 반환합니다.
        rewards = []
        for row in range(self.reward_table.rowCount()):
            reward_type = self.reward_table.item(row, 1).text()
            reward_count = self.reward_table.item(row, 2).text()
            reward_info_id = self.reward_table.item(row, 3).text()
            item_bind = self.reward_table.item(row, 4).text()
            event_item_period_info_id = self.reward_table.item(row, 5).text()
            reward_data = {
                'reward_type': reward_type,
                'reward_count': reward_count,
                'reward_info_id': reward_info_id,
                'item_bind': item_bind,
                'event_item_period_info_id': event_item_period_info_id
            }
            rewards.append(reward_data)

        return {'tab_name': self.main_window.tab_widget.tabText(self.main_window.tab_widget.indexOf(self)),
                'rewards': rewards}

    def from_json(self, data):
        # JSON 데이터에서 보상 정보를 가져와서 reward_table에 추가합니다.
        rewards = data.get('rewards', [])
        for reward in rewards:
            reward_type = reward.get('reward_type', '')
            reward_count = reward.get('reward_count', '')
            reward_info_id = reward.get('reward_info_id', '')
            item_bind = reward.get('item_bind', '')
            event_item_period_info_id = reward.get('event_item_period_info_id', '')
            self.add_reward_to_table(reward_type, reward_count, reward_info_id, item_bind, event_item_period_info_id)

            tab_name = data.get('tab_name', '보상')
            self.main_window.tab_widget.setTabText(self.main_window.tab_widget.indexOf(self), tab_name)
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('보상 지급 엑셀파일 제작')
        self.setGeometry(0, 0, 600, 800)
        self.layout = QVBoxLayout()
        self.central_widget = QWidget()
        self.central_widget.setLayout(self.layout)
        self.setCentralWidget(self.central_widget)
        self.tab_widget = QTabWidget()
        self.layout.addWidget(self.tab_widget)
        self.load_data()

        self.close_button = QPushButton('종료', self)
        self.close_button.clicked.connect(self.close_app)
        self.layout.addWidget(self.close_button)

        # 탭 더블 클릭 시 이름을 변경하는 함수 연결
        self.tab_widget.tabBarDoubleClicked.connect(self.on_tab_double_click)

    def add_tab(self):
        # 탭 추가 함수
        tab_index_count = self.tab_widget.count()
        tab_index = self.tab_widget.addTab(MainTab(self), f'보상 {tab_index_count + 1}')
        self.tab_widget.setCurrentIndex(tab_index)


    def close_tab(self, index):
        # 탭 삭제 함수
        if self.tab_widget.count() > 1:
            self.tab_widget.removeTab(index)
        else:
            QMessageBox.warning(self, '경고', '마지막 탭은 닫을 수 없습니다.')

    def keyPressEvent(self, event):
        # 키 입력 이벤트 함수
        if event.modifiers() == Qt.ControlModifier and event.key() == Qt.Key_T:
            self.add_tab()
        elif event.modifiers() == Qt.ControlModifier and event.key() == Qt.Key_W:
            current_index = self.tab_widget.currentIndex()
            self.close_tab(current_index)

    def load_data(self):
        # data 파일 불러오기 함수
        try:
            with open('data.json', 'r', encoding='utf-8') as file:
                data = json.load(file)
                tab_data_list = data.get('tabs', [])
                for tab_data in tab_data_list:
                    tab = MainTab(self)
                    tab.from_json(tab_data)
                    tab_index = self.tab_widget.addTab(tab, tab_data.get('tab_name', '보상'))
                    self.tab_widget.setCurrentIndex(tab_index)

        except Exception as e:
            QMessageBox.warning(self, '오류', f'불러올 데이터가 없거나, 데이터 로드 중 오류가 발생했습니다: {str(e)}')
            self.add_tab()

    def save_data(self):
        # data 파일 저장 함수
        tab_data_list = []
        for tab_index in range(self.tab_widget.count()):
            tab = self.tab_widget.widget(tab_index)
            tab_data = tab.to_json()
            tab_data_list.append(tab_data)

        data = {'tabs': tab_data_list}
        with open('data.json', 'w', encoding='utf-8') as file:
            json.dump(data, file, ensure_ascii=False, indent=2)

    def closeEvent(self, event):
        # 앱 종료 이벤트 함수
        self.save_data()
        event.accept()

    def close_app(self):
        # 앱 종료 버튼 클릭시 호출 함수
        self.save_data()
        app.quit()
        logger.info('앱 종료')

    def on_tab_double_click(self, index):
        # 탭 이름 변경 함수
        current_tab_text = self.tab_widget.tabText(index)
        new_tab_text, ok = QInputDialog.getText(self, '탭 이름 변경', '새로운 탭 이름을 입력하세요:', text=current_tab_text)

        if ok and new_tab_text:
            self.tab_widget.setTabText(index, new_tab_text)
        logger.info(f'탭 이름 변경 기존 탭 이름 : {current_tab_text}, 변경 탭 이름 : {new_tab_text}')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    apply_stylesheet(app, theme='dark_teal.xml')
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())