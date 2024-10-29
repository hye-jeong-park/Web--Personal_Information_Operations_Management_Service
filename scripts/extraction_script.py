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


def navigate_to_search_page(driver: webdriver.Chrome) -> bool:
    """
    검색 페이지로 이동하는 함수
    """
    try:
        # 검색 페이지 URL로 이동
        driver.get('https://gw.com2us.com/emate_appro/appro_complete_2024_link.nsf/wfmViaView?readform&viewname=view055&vctype=a')
        # 검색창이 로드될 때까지 대기
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'searchtext')))
        return True
    except Exception as e:
        logging.error(f"검색 페이지 이동 중 오류 발생: {e}")
        traceback.print_exc()
        return False


def search_documents(driver: webdriver.Chrome) -> bool:
    """
    '개인정보 추출 신청서' 검색을 수행하는 함수
    """
    try:
        # 검색어 입력
        search_input = driver.find_element(By.ID, 'searchtext')
        search_input.clear()
        search_input.send_keys('개인정보 추출 신청서')
        
        # 검색 버튼 클릭
        search_button = driver.find_element(By.XPATH, '//img[@class="inbtn" and contains(@src, "btn_search_board.gif")]')
        search_button.click()
        time.sleep(5)
        return True
    except Exception as e:
        logging.error(f"문서 검색 중 오류 발생: {e}")
        traceback.print_exc()
        return False


def extract_post_data(driver: webdriver.Chrome, post: webdriver.remote.webelement.WebElement, index: int) -> Optional[Dict]:
    """
    개별 게시글의 데이터를 추출하는 함수
    """
    try:
        # 기본 정보 추출
        tds = post.find_elements(By.TAG_NAME, 'td')
        
        # 결재일 추출
        결재일_text = tds[5].text.strip()
        년, 월, 일 = 결재일_text.split('-')
        월 = str(int(월))
        일 = str(int(일))
        
        # 신청자 추출
        신청자 = tds[4].find_element(By.TAG_NAME, 'span').text.strip()

        # 게시글 상세 페이지 열기
        driver.execute_script("arguments[0].scrollIntoView();", post)
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(post))
        post.click()

        # 새 창으로 전환
        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
        driver.switch_to.window(driver.window_handles[-1])

        # 상세 페이지 데이터 추출
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'AppLineArea')))
        
        # 문서 종류 확인
        h2_element = driver.find_element(By.CSS_SELECTOR, '#AppLineArea h2')
        if h2_element.text.strip() != '개인정보 추출 신청서':
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
            return None

        # 상세 정보 추출
        법인명_elements = driver.find_elements(By.ID, 'titleLabel')
        법인명 = 법인명_elements[0].text.strip() if 법인명_elements else ''

        문서번호_elements = driver.find_elements(By.XPATH, '//th[contains(text(),"문서번호")]/following-sibling::td[1]')
        문서번호 = 문서번호_elements[0].text.strip() if 문서번호_elements else ''

        제목_elements = driver.find_elements(By.CSS_SELECTOR, 'td.approval_text')
        제목 = 제목_elements[0].text.strip().replace(법인명, '').strip() if 제목_elements else ''

        합의담당자_elements = driver.find_elements(By.XPATH, '//th[text()="합의선"]/following::tr[@class="name"][1]/td[@class="td_point"]')
        합의담당자 = 합의담당자_elements[0].text.strip() if 합의담당자_elements else ''

        # 추출 데이터 구성
        data = {
            '결재일': 결재일_text,
            '년': 년,
            '월': 월,
            '일': 일,
            '주차': '',
            '법인명': 법인명,
            '문서번호': 문서번호,
            '제목': 제목,
            '업무 유형': '',
            '추출 위치': '',
            '담당 부서': '',
            '신청자': 신청자,
            '합의 담당자': 합의담당자,
            '링크': driver.current_url,
            '진행 구분': ''
        }

        logging.info(f"게시글 {index}: 데이터 추출 완료")
        return data

    except Exception as e:
        logging.error(f"게시글 {index}: 데이터 추출 중 오류 발생: {e}")
        traceback.print_exc()
        return None
    finally:
        # 창 정리 및 원래 창으로 복귀
        try:
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
            time.sleep(2)
        except Exception as e:
            logging.error(f"창 전환 중 오류 발생: {e}")
            traceback.print_exc()