import pandas as pd
import numpy as np
import locale, glob, math, time, random, warnings, datetime, requests, ssl
import requests, json, time, re, os
import xlwings as xw
from webdriver_manager.chrome import ChromeDriverManager

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup as bs
os.environ['WDM_SSL_VERIFY'] = '0' ## ChromeDriverManager().install() 실행 시 verify = False 옵션

############################################################################
## 이메일 발송 설정: notebooks/.env 에서 읽어옴
## 비밀번호를 이 파일이나 노트북에 직접 적지 말 것 (.env 만 고치면 됨)
############################################################################
import sys
from pathlib import Path
_NOTEBOOKS_ROOT = Path(__file__).resolve().parents[2]   ## notebooks 폴더
if str(_NOTEBOOKS_ROOT) not in sys.path:
    sys.path.insert(0, str(_NOTEBOOKS_ROOT))
from mail_env import load_mail_config, require_password

## display() 는 노트북에서만 기본 제공됨. 아래 import 로 일반 python 실행도 가능해짐
from IPython.display import display

_cfg = load_mail_config("RECV_ADDRS_LAW_REVISION")
SEND_ADDR       = _cfg.send_addr
EMAIL_PASSWORD  = _cfg.password
SMTP_SERVER     = _cfg.smtp_server
SMTP_PORT       = _cfg.smtp_port
RECV_ADDRS      = _cfg.recv_addrs
DOWNLOAD_FOLDER = _cfg.download_folder


## Excel 창을 띄울지 여부. 자동 실행(run_check.py) 에서는 False 로 끈다.
EXCEL_VISIBLE = True

def set_excel_visible(visible):
    global EXCEL_VISIBLE
    EXCEL_VISIBLE = visible

def excel_is_visible():
    return EXCEL_VISIBLE

def open_excel_app():
    """서식 작업용 Excel 인스턴스. EXCEL_VISIBLE 설정을 따른다."""
    return xw.App(visible=EXCEL_VISIBLE)

## 이 도구가 실제로 내려받는 문서 확장자. 삭제 대상을 이 형식으로만 제한함
DOWNLOAD_FILE_EXTS = ('.hwp', '.hwpx', '.pdf', '.doc', '.docx')

## 법규리스트 엑셀파일 강제로 닫기
def close_law_list_excel():
    # 대상 파일 경로 (절대경로로 지정하는 것이 안전함)
    file_name = "법규리스트.xlsx"

    # 열려있는 모든 엑셀 앱의 모든 워크북 확인
    for app in xw.apps:
        for book in app.books:            
            if book.name.find(file_name)>-1:            
                book.close()
        if len(app.books)==0:
            app.quit()

## 법령과 행정규칙의 이름과 생성일 칼럼명을 얻기 위한 함수
def get_column_name(site_category, get_address=False):
    name_col_dict = {
        'law':'법령명', 'reg':'행정규칙명'
    }
    date_col_dict = {
        'law':'공포일자', 'reg':'발령일자'
    }    
    date_col = date_col_dict[site_category]
    name_col = name_col_dict[site_category]       
    return name_col, date_col

## 법령과 행정규칙 메인사이트 주소 찾기. 모든 것의 출발점임
def get_url(site_category):
    url_dict = {
        'law': "https://www.law.go.kr/lsSc.do?menuId=1&subMenuId=23&tabMenuId=121&query=",
        'reg': "https://www.law.go.kr/admRulSc.do?menuId=5&subMenuId=45&tabMenuId=203&query="
    }
    return url_dict[site_category]

def get_browser():
    return webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

def remove_dup_files():  ## 시작전에 다운로드 폴더의 중복 파일을 제거함
    """이전 실행에서 받아둔 법령 파일을 지움.

    [주의] 예전 코드는 법규리스트의 법령명 163개 각각에 대해 다운로드 폴더 전체를
    '부분 일치'로 훑어 지웠다. 그래서 '개인정보보호법 스터디자료.hwp' 처럼
    이름이 겹치는 개인 파일까지 조용히 삭제될 수 있었다.

    지금은 아래 세 가지로 좁혔다.
      (1) 이 도구가 실제로 내려받는 문서 확장자만 대상으로 함
      (2) 파일명이 법령명으로 '시작'하는 경우만 대상으로 함
      (3) 무엇을 지웠는지 출력함 (조용한 삭제 금지)
    """
    law_df = pd.read_excel("./database/법규리스트.xlsx")
    keywords = law_df.법령명.values.tolist()

    removed = []
    for filename in os.listdir(DOWNLOAD_FOLDER):
        file_path = os.path.join(DOWNLOAD_FOLDER, filename)
        if not os.path.isfile(file_path):
            continue
        file_name, file_ext = os.path.splitext(filename)
        if file_ext.lower() not in DOWNLOAD_FILE_EXTS:  ## 이 도구가 받는 형식이 아니면 건드리지 않음
            continue
        normalized = re.sub(r'[^가-힣]', '', file_name)
        if not normalized:
            continue
        if any(normalized.startswith(keyword) for keyword in keywords):
            os.remove(file_path)
            removed.append(filename)

    if removed:
        print(f"  기존 다운로드 파일 {len(removed)}건 삭제: "
              + ", ".join(removed[:3]) + (" 외" if len(removed) > 3 else ""))


## 메일에 첨부해야할 파일 리스트업
def find_law_files(law_name, since=None):
    """메일에 첨부할 파일 목록을 만듦.

    since (time.time() 값) 를 주면 그 시각 이후에 저장된 파일만 대상으로 한다.
    다운로드 폴더에 원래 있던 개인 파일이 이름이 비슷하다는 이유로
    메일에 잘못 첨부되는 것을 막기 위한 장치이므로, 되도록 항상 넘길 것.
    """
    file_list = []
    for filename in os.listdir(DOWNLOAD_FOLDER):
        file_path = os.path.join(DOWNLOAD_FOLDER, filename)
        if not os.path.isfile(file_path):
            continue
        if since is not None and os.path.getmtime(file_path) < since:
            continue  ## 이번 실행에서 받은 파일이 아님
        file_name, file_ext = os.path.splitext(filename)
        file_name = re.sub(r'[^가-힣]', '', file_name)
        if file_name.find(law_name)>-1: ## LIKE 검색을함. 그래서 파일첨부시 주의할 필요
            file_list.append(file_path)
    return file_list


## 법령-최신법령/행정규칙 초기 화면으로 이동
def move_to_home(browser, site_category):
    url = get_url(site_category)    
    browser.get(url)
    WebDriverWait(browser, 10).until(
        lambda browser: browser.execute_script('return document.readyState') == 'complete'
    )

## 법령 마지막 페이지 넘버를 찾는 함수
def get_last_page_number(browser, site_category):
    ## 마지막 페이지로 이동
    for tag in browser.find_elements(By.TAG_NAME, 'img'):
        property_alt = tag.get_property('alt').strip()

        if property_alt == '마지막으로':
            try:
                tag.click()
                break
            except:
                continue
    time.sleep(1)
    for tag in browser.find_elements(By.CLASS_NAME, 'on'):
        ## 1~5자리 숫자로만 구성된 것. 즉 마지막 페이지 번호
        if re.search(r'^\d{1,5}$', tag.text):
            last_page_number = int(tag.text)
            break
    move_to_home(browser, site_category)
    time.sleep(0.5)
    return last_page_number

### 법령 데이터 업데이트 할때 크롤링할 페이지 리스트를 생성
def get_page_range(last_page_number):
    page_ranges = []
    table_df_list = []

    for i in range(1, last_page_number, 5):  ## 하나의 페이지엔 5개의 세부 페이지가 존재
        if i + 4 < last_page_number:
            page_ranges.append((i, i + 4))
        else:
            page_ranges.append((i, last_page_number))
    return page_ranges

## 해당페이지 번호로 이동
def move_to_page(browser, page_num):
    click_yn = 0
    for page_tag in browser.find_elements(By.CLASS_NAME, 'paging'):
        if page_tag.text:
            for tag in page_tag.find_elements(By.TAG_NAME, 'li'):
                if int(tag.text) == page_num:
                    tag.click()
                    click_yn = 1
                    break
        if click_yn:
            break
    WebDriverWait(browser, 10).until(
        lambda browser: browser.execute_script('return document.readyState') == 'complete'
    )

## 다음 페이지로 이동
def click_next_page(browser):
    for tag in browser.find_elements(By.TAG_NAME, 'img'):
        property_alt = tag.get_property('alt').strip()

        if property_alt == '다음으로':
            try:
                tag.click()
                break
            except:
                continue

            ## 페이지의 테이블을 dataframe으로 반환(속도를 위해 beautifulsoup으로 구현)

## 현재 페이지의 법령정보 테이블을 크롤링함
def get_page_table_info(page_source):
    table_cols = []
    table_tag = page_source.find('table')

    for tag in table_tag.find_all('th', attrs={'scope': 'col'}):
        text = tag.text.strip()
        if text:
            table_cols.append(text)

    table_df = pd.DataFrame(columns=table_cols)

    ## dataframe에 들어갈 데이터 수집
    for tr_idx, tr in enumerate(table_tag.find_all('tr')):
        for td_idx, td in enumerate(tr.find_all('td')):
            text = td.text.strip()
            if text:
                table_df.loc[tr_idx, table_cols[td_idx]] = text
    return table_df