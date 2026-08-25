import pandas as pd
import numpy as np
import locale, glob, math, time, random, warnings, datetime, requests, ssl, smtplib
import requests, json, time, re, os
import xlwings as xw
from webdriver_manager.chrome import ChromeDriverManager

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.encoders import encode_base64

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup as bs
os.environ['WDM_SSL_VERIFY'] = '0' ## ChromeDriverManager().install() 실행 시 verify = False 옵션
warnings.filterwarnings('ignore')

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

_cfg = load_mail_config("RECV_ADDRS_FSC_INTERPRETATION")
SEND_ADDR       = _cfg.send_addr
EMAIL_PASSWORD  = _cfg.password
SMTP_SERVER     = _cfg.smtp_server
SMTP_PORT       = _cfg.smtp_port
RECV_ADDRS      = _cfg.recv_addrs
DOWNLOAD_FOLDER = _cfg.download_folder


## 다운로드 받으려는 PDF 파일이 있을 경우 먼저 삭제
def remove_dup_files(download_folder, remove_file):    
    folder_path = download_folder
    for file in os.listdir(folder_path):
        fn, _ = os.path.splitext(file)    
        if fn.find(remove_file)>-1:        
            os.remove(os.path.join(folder_path, file))

def get_browser():
    return webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

def move_to_home(browser):
    browser.get("https://better.fsc.go.kr/fsc_new/replyCase/PastReplyList.do?stNo=11&muNo=171&muGpNo=75")

def get_num_tag(browser):
    num_tag = False
    for tag in browser.find_elements(By.TAG_NAME, 'div'):
        if tag.get_attribute('class')=="dataTables_paginate paging_full_numbers":
            num_tag = tag
    return num_tag if num_tag else None

def move_to_first_page(browser):
    move_to_home(browser)

    num_tag = get_num_tag(browser)   
    
    for a_tag in num_tag.find_elements(By.TAG_NAME, 'a'):        
        if re.search('first', a_tag.get_property('id')):
            a_tag.click()
    time.sleep(0.3)

## 법령 해석 리스트의 가장 마지막 페이지 번호 호출
def get_last_page_num(browser):
    move_to_home(browser)
    time.sleep(0.1)
    num_tag = get_num_tag(browser)   
    time.sleep(0.1)    
    num_list = []
    for a_tag in num_tag.find_elements(By.TAG_NAME, 'a'):        
        page_num = a_tag.text.strip()
        if page_num:
            num_list.append(page_num) 
    move_to_first_page(browser)
    return int(num_list[-1]) ## 보이는 숫자 중 가장 마지막 숫자가 마지막 페이지번호
    
## 법령해석 리스트 페이지의 페이지 정보와 표 본문 태그를 반환
def get_table_data(browser):
    thead_tag = browser.find_element(By.TAG_NAME, 'thead')
    tbody_tag = browser.find_element(By.TAG_NAME, 'tbody')
    
    cols = []
    for tr in thead_tag.find_elements(By.TAG_NAME, 'th'):
        cols.append(tr.text.strip())
    
    table_df = pd.DataFrame(columns=cols)
    
    for row_idx, tr in enumerate(tbody_tag.find_elements(By.TAG_NAME, 'tr')):    
        for col_idx, td in enumerate(tr.find_elements(By.TAG_NAME, 'td')):
            table_df.loc[row_idx, cols[col_idx]] = td.text.strip()
    
    return table_df, tbody_tag

## 해당 법령해석 상세 페이지의 정보를 가져옴.
def get_page_info(browser, download_folder=None):
    if download_folder is None:  ## 지정하지 않으면 .env 의 DOWNLOAD_FOLDER 사용
        download_folder = DOWNLOAD_FOLDER
    table_tag = browser.find_elements(By.TAG_NAME, 'table')[1] ## 세부 페이지에서 가져올 표는 두번째 표(회신)임
    tbody_tag = table_tag.find_element(By.TAG_NAME, 'tbody')
    col_name = tbody_tag.find_elements(By.TAG_NAME, 'tr')[0].text
    page_df = pd.DataFrame(columns=[col_name])
    
    for tr in tbody_tag.find_elements(By.TAG_NAME, 'tr')[1:]:    
        idx_name = tr.find_element(By.TAG_NAME, 'th').text.strip()
        td_text = tr.find_element(By.TAG_NAME, 'td').text.strip()
        if idx_name=='첨부파일': ## 첨부파일은 다운로드 받음
            remove_dup_files(download_folder, os.path.splitext(td_text)[0])
            tr.find_element(By.TAG_NAME, 'td').find_element(By.TAG_NAME, 'a').click()
            
        page_df.loc[idx_name, col_name] = td_text
    return page_df

## 최종적으로 이메일 발송
def send_mail(page_info, download_folder=None):
    require_password(EMAIL_PASSWORD)  ## 비밀번호 미설정 시 발송 전에 중단
    if download_folder is None:  ## 지정하지 않으면 .env 의 DOWNLOAD_FOLDER 사용
        download_folder = DOWNLOAD_FOLDER
    update_date = pd.to_datetime(page_info.loc['회신일'].values[0]).strftime("%Y%m%d")
    law_name = page_info.columns[0]
    law_name = re.sub(r'[\\/*?:"<>|]', '', law_name)

    page_info_file = f"./update_list/{update_date}_{law_name}.xlsx"
    page_info.to_excel(page_info_file)
    app = xw.App(visible=True)
    wb = app.books.open(page_info_file)
    sht = wb.sheets[0]        
    wb.save()
    wb.close()
    app.quit()
    
    mail_title = f"(법령해석포털)_{update_date}_{law_name}"
    mail_body = "본문의 내용은 첨부한 엑셀파일을 참고해주세요." 
    # 메일 객체 생성 및 로그인
    mail_server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    mail_server.ehlo()
    mail_server.starttls()
    mail_server.ehlo()
    mail_server.login(SEND_ADDR, EMAIL_PASSWORD)
    
    for recv_addr in RECV_ADDRS:
        # 제목, 본문 작성
        msg = MIMEMultipart()
        msg['From'] = SEND_ADDR
        msg['To'] = recv_addr
        msg['Subject'] = mail_title
        msg.attach(MIMEText(mail_body, _charset='utf-8'))           
        
        # 파일첨부
        files = [page_info_file, os.path.join(download_folder, page_info.loc['첨부파일'].values[0])]
        for file in files:
            if file:                
                part = MIMEBase('application', "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                part.set_payload(open(file, "rb").read())
                encode_base64(part)
                part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file))
                msg.attach(part)

        ## 최종 발송
        mail_server.sendmail(SEND_ADDR, msg['To'].split(','), msg.as_string())

    mail_server.quit()  ## SMTP 연결 정리

## 각 항목들을 클릭하고 페이지 정보를 수집한 후 메일까지 발송
def get_page_laws_info(browser, table_df, download_folder=None):
    if download_folder is None:  ## 지정하지 않으면 .env 의 DOWNLOAD_FOLDER 사용
        download_folder = DOWNLOAD_FOLDER
    ## table_df에 있는 항목만 가져옴
    get_info_nums = table_df.번호.values.tolist()    
    for get_info_num in get_info_nums:
        find_continue = True
        _, tbody_tag = get_table_data(browser)
        for tr in tbody_tag.find_elements(By.TAG_NAME, 'tr'):
            if find_continue==False:
                break
            td = tr.find_element(By.TAG_NAME, 'td').text.strip()        
            if td==get_info_num:
                for td_tag in tr.find_elements(By.TAG_NAME, 'td'): ## 실제 표의 내용들 확인하면서                
                    if td_tag.text.strip()==table_df[table_df.번호==get_info_num]['제목'].values[0]: ## 제목과 같으면 클릭
                        while True:
                            try:
                                ## 해당 법령해석 상세피이지로 이동
                                a_tag = td_tag.find_element(By.TAG_NAME, 'a')
                                browser.execute_script("arguments[0].click();", a_tag)                                
                                break
                            except:
                                time.sleep(0.01)                                
                        find_continue = False
                        time.sleep(0.3)
                        page_info = get_page_info(browser, download_folder) ## 상세 페이지의                         
                        time.sleep(0.2)
                        browser.back()
                        send_mail(page_info, download_folder)
                        break        
    
        time.sleep(0.2)
    

## 업데이트된 법령이 있는지만 체크
def update_check(browser, today):
    last_page_num = get_last_page_num(browser)
    table_df, tbody_tag = get_table_data(browser)
    table_df['등록일'] = table_df['등록일'].map(lambda x: pd.to_datetime(x).strftime("%y%m%d"))
    
    table_df = table_df[table_df.등록일>=today]
    if len(table_df)>0:
        display(table_df)
        print("새로 올라온 법령해석 정보가 있습니다. 다음 셀을 실행해주세요.")
    else:
        print('새로 올라온 법령해석 정보가 없습니다.')