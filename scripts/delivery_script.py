import re
import sys
import time
import traceback
import logging
from typing import Tuple, Optional, List, Dict

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook

# 크롤링할 최대 게시글 수 설정
CRAWL_LIMIT = 21
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
        driver.get('https://gw.com2us.com/')
        username_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'Username'))
        )
        password_input = driver.find_element(By.ID, 'Password')

        username_input.send_keys(username)
        password_input.send_keys(password)
        driver.find_element(By.CLASS_NAME, 'btnLogin').click()

        WebDriverWait(driver, 20).until(EC.url_changes('https://gw.com2us.com/'))
        if 'login' in driver.current_url.lower():
            logging.error("로그인에 실패하였습니다.")
            return False
        return True
    except Exception as e:
        logging.error(f"로그인 중 오류 발생: {e}")
        traceback.print_exc()
        return False


def navigate_to_target_page(driver: webdriver.Chrome) -> bool:
    """
    개인정보 파일 전송 페이지로 이동
    """
    try:
        driver.get('https://gw.com2us.com/emate_app/00001/bbs/b2307140306.nsf/view?readform&viewname=view01')
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'tr[class*="dhx_skyblue"]'))
        )
        logging.info(f"페이지 이동 후 현재 URL: {driver.current_url}")
        return True
    except Exception as e:
        logging.error(f"타겟 페이지로 이동 중 오류 발생: {e}")
        traceback.print_exc()
        return False


def fetch_posts(driver: webdriver.Chrome) -> List[webdriver.remote.webelement.WebElement]:
    """
    현재 페이지의 게시글 목록을 가져옵니다.
    """
    try:
        posts = driver.find_elements(By.CSS_SELECTOR, 'tr[class*="dhx_skyblue"]')
        total_posts = len(posts)
        logging.info(f"총 게시글 수: {total_posts}")
        return posts
    except Exception as e:
        logging.error("게시글 목록을 가져오는 중 오류 발생.")
        logging.error(e)
        traceback.print_exc()
        return []


def go_to_page(driver: webdriver.Chrome, page_number: int) -> bool:
    """
    주어진 페이지 번호로 이동합니다.
    """
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'pagingNav'))
        )

        current_page_element = driver.find_element(By.CSS_SELECTOR, 'div#pagingNav strong.cur_num')
        current_page = int(current_page_element.text.strip())
        if current_page == page_number:
            return True

        page_links = driver.find_elements(By.XPATH, f'//div[@id="pagingNav"]//a[@class="num_box"]')
        page_link = None
        for link in page_links:
            if link.text.strip() == str(page_number):
                page_link = link
                break

        if page_link:
            page_link.click()
        else:
            logging.info(f"페이지 번호 {page_number}를 찾을 수 없습니다.")
            return False

        time.sleep(2)
        WebDriverWait(driver, 10).until(
            EC.text_to_be_present_in_element((By.CSS_SELECTOR, 'div#pagingNav strong.cur_num'), str(page_number))
        )
        logging.info(f"{page_number} 페이지로 이동 완료")
        return True
    except Exception as e:
        logging.error(f"{page_number} 페이지로 이동 중 오류 발생: {e}")
        traceback.print_exc()
        return False


def extract_corporate_name(full_text: str) -> str:
    """
    법인명 추출: "게임사업3본부 K사업팀 / 홍길동님" 중 "게임사업3본부"만 추출
    """
    if '/' in full_text:
        return full_text.split('/')[0].strip().split()[0]
    return full_text.strip().split()[0]


def extract_file_info(file_info: str) -> Tuple[str, str]:
    """
    파일형식 및 파일 용량 추출
    """
    file_match = re.match(r'(.+?)\s*(?:&|[(])\s*([\d,\.]+\s*[KMGT]?B)', file_info, re.IGNORECASE)
    if file_match:
        filename_part = file_match.group(1).strip()
        size_part = file_match.group(2).strip()
    else:
        filename_part = file_info.strip()
        size_match = re.search(r'([\d,\.]+\s*[KMGT]?B)', filename_part, re.IGNORECASE)
        if size_match:
            size_part = size_match.group(1).strip()
            filename_part = filename_part.replace(size_part, '').strip()
        else:
            size_part = ''

    file_type = ''
    if '.zip' in filename_part.lower():
        file_type = 'Zip'
    elif '.xlsx' in filename_part.lower():
        file_type = 'Excel'

    size_match = re.match(r'([\d,\.]+)\s*([KMGT]?B)', size_part, re.IGNORECASE)
    if size_match:
        size_numeric = size_match.group(1).replace(',', '')
        size_unit = size_match.group(2).upper()
        file_size = f"{size_numeric} {size_unit}"
    else:
        file_size = size_part

    return file_type, file_size


def find_section_text(driver: webdriver.Chrome, section_titles: List[str]) -> Optional[str]:
    """
    특정 섹션의 제목을 기반으로 해당 섹션의 내용을 추출하는 함수
    """
    try:
        tr_elements = driver.find_elements(By.XPATH, '//table//tr')
        for tr in tr_elements:
            try:
                td_elements = tr.find_elements(By.TAG_NAME, 'td')
                if len(td_elements) >= 2:
                    header_text = ''.join([span.text.strip() for span in td_elements[0].find_elements(By.TAG_NAME, 'span')])

                    for section_title in section_titles:
                        if section_title in header_text:
                            return td_elements[1].text.strip()
            except Exception:
                continue
        return None
    except Exception as e:
        logging.error(f"find_section_text 오류: {e}")
        return None


def extract_attachment_info(driver: webdriver.Chrome) -> Tuple[str, str]:
    """
    메인 문서 내의 첨부파일 정보를 추출하는 함수
    """
    파일형식, 파일용량 = '', ''

    try:
        attm_read_div = driver.find_element(By.ID, 'attmRead')
        logging.info("첨부파일 div 찾음: attmRead")

        try:
            size_text = attm_read_div.find_element(By.XPATH, './/span[@class="attm-size"]').text.strip()
            size_match = re.match(r'([\d,\.]+)\s*([KMGT]?B)', size_text, re.IGNORECASE)
            if size_match:
                size_numeric = size_match.group(1).replace(',', '')
                size_unit = size_match.group(2).upper()
                파일용량 = f"{size_numeric} {size_unit}"
            else:
                파일용량 = size_text
            logging.info(f"파일용량 추출: {파일용량}")
        except Exception as e:
            logging.warning(f"파일용량 추출 중 오류 발생: {e}")

        try:
            filename = attm_read_div.find_element(By.XPATH, './/ul[contains(@class, "attm-list")]/li/a/strong').text.strip()
            if '.zip' in filename.lower():
                파일형식 = 'Zip'
            elif '.xlsx' in filename.lower():
                파일형식 = 'Excel'
            logging.info(f"파일형식 추출: {파일형식}")
        except Exception as e:
            logging.warning(f"파일형식 추출 중 오류 발생: {e}")
            파일형식 = ''
    except Exception as e:
        logging.warning(f"attmRead를 찾을 수 없음: {e}")

    if not 파일형식 and not 파일용량:
        try:
            iframe = driver.find_element(By.ID, 'ifa_form')
            driver.switch_to.frame(iframe)
            logging.info("iframe으로 전환하여 파일 정보 추출 시도")
            file_text = find_section_text(driver, ['파밀명 및 용량 (KB)', '파일명 및 용량 (KB)'])
            if file_text:
                logging.info(f"iframe 내에서 파일 정보 추출 시작: {file_text}")
                파일형식, 파일용량 = extract_file_info(file_text)
                logging.info(f"iframe 내에서 파일 정보 추출 완료: {파일형식}, {파일용량}")
            else:
                logging.warning("iframe 내에서 파일 정보 섹션을 찾을 수 없습니다.")
            driver.switch_to.default_content()
        except Exception as e:
            logging.error(f"iframe에서 파일 정보 추출 중 오류 발생: {e}")
            driver.switch_to.default_content()

    return 파일형식, 파일용량