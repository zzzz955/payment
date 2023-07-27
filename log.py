import datetime
import logging
import os

def setup_logger():
    # 로그 파일 설정
    current_datetime = datetime.datetime.now()
    log_directory = 'logs'
    log_filename = f'log_{current_datetime.strftime("%Y%m%d_%H%M%S")}.txt'
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
