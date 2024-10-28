import sys
import time
import traceback
import logging
from typing import Optional, List, Dict

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook

# 크롤링할 최대 게시글 수 설정
MAX_POSTS = 20
EXCEL_FILE = r'C:\Users\PHJ\output\개인정보 운영대장.xlsx'  # 엑셀 파일 경로
WORKSHEET_NAME = '개인정보 추출 및 이용 관리'

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def initialize_webdriver() -> webdriver.Chrome:
    """
    웹드라이버 초기화
    """
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)
    return driver


def login(driver: webdriver.Chrome, username: str, password: str) -> bool:
    """
    로그인 처리
    """
    try:
        # 로그인 페이지로 이동
        driver.get('https://gw.com2us.com/')
        
        # 로그인 폼 요소 찾기
        username_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'Username'))
        )
        password_input = driver.find_element(By.ID, 'Password')

        # 로그인 정보 입력 및 제출
        username_input.send_keys(username)
        password_input.send_keys(password)
        driver.find_element(By.CLASS_NAME, 'btnLogin').click()

        # 로그인 성공 여부 확인
        WebDriverWait(driver, 30).until(EC.url_changes('https://gw.com2us.com/'))
        if 'login' in driver.current_url.lower():
            logging.error("로그인에 실패하였습니다.")
            return False
        return True
    except Exception as e:
        logging.error(f"로그인 중 오류 발생: {e}")
        traceback.print_exc()
        return False