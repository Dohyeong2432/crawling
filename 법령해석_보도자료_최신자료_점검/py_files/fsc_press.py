import pandas as pd
import numpy as np
import locale, glob, math, time, random, warnings, datetime, requests, ssl, smtplib
import requests, json, time, re, os
import xlwings as xw
from py_files.common_functions import *

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.encoders import encode_base64

from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup as bs
os.environ['WDM_SSL_VERIFY'] = '0' ## ChromeDriverManager().install() 실행 시 verify = False 옵션
warnings.filterwarnings('ignore')

## 환경변수 설정 값은 common_functions.py 에서 한 번만 읽어와 공유함
## (from py_files.common_functions import * 로 SEND_ADDR/DOWNLOAD_FOLDER 등이 들어옴)


## 해당 페이지의 테이블 데이터 긁어오기
def get_table_data(browser):
    list_tag = browser.find_element(By.CLASS_NAME, 'board-list-wrap')
    li_tags = list_tag.find_elements(By.TAG_NAME, 'li')
    table_df = pd.DataFrame(columns=['제목', '등록일'])
    row_idx = 0
    for li_tag in li_tags:
        date = li_tag.find_element(By.CLASS_NAME, 'day').text.strip()
        date = pd.to_datetime(date).strftime("%y%m%d")
        subject = li_tag.find_element(By.CLASS_NAME, 'cont').find_element(By.CLASS_NAME, 'subject').text.strip()
        table_df.loc[row_idx, '제목'] = subject
        table_df.loc[row_idx, '등록일'] = date
        row_idx += 1
    return table_df

## 새로운 보도자료가 나와있는지 확인
def update_check(browser, today):
    move_to_home(browser, 'fsc_press')
    table_df = get_table_data(browser)
    table_df = table_df[table_df.등록일>=today]
    if len(table_df) > 0:        
        print("(금융위 보도자료) 새로 올라온 자료가 있습니다. 아래 표를 참고해주세요.")
        display(table_df)
    else:
        print("(금융위 보도자료) 새로 올라온 자료는 없습니다.")

## 해당 페이지 번호로 이동
def move_to_page(browser, page_num):
    page_area = browser.find_element(By.CLASS_NAME, 'paginate')
    for a_tag in page_area.find_elements(By.TAG_NAME, 'a'):
        if a_tag.text.strip()==str(page_num):
            a_tag.click()
            break    

## 현재 페이지의 테이블 정보를 확인하고, 최신정보면 파일을 다운로드해 이메일까지 발송함
def download_send_mail(browser, today, title_keywords=None):
    sended_mail_subjects = []
    sended_mail_dates = []
    list_tag = browser.find_element(By.CLASS_NAME, 'board-list-wrap')
    li_tags = list_tag.find_elements(By.TAG_NAME, 'li')
    for li_tag in li_tags:
        count = li_tag.find_element(By.CLASS_NAME, 'count').text.strip()        
        date = li_tag.find_element(By.CLASS_NAME, 'day').text.strip()
        date = pd.to_datetime(date).strftime("%y%m%d")
        subject = li_tag.find_element(By.CLASS_NAME, 'subject').text.strip()
        mail_title = f"(금융위보도자료)_{date}_{subject}" ## 메일 제목 양식 수정
        
        ## 최신 보도자료면 다운로드함
        download_files = []
        if date>=today and subject_matches(subject, title_keywords):
            for file_tag in li_tag.find_elements(By.CLASS_NAME, 'file-list'):
                file_type = file_tag.text.split('.')[-1]
                if re.search('hwp', file_type) and re.search('hwpx', file_type) is None:                
                    for tag in file_tag.find_elements(By.TAG_NAME, 'span'):
                        file_name = os.path.splitext(file_tag.text)[0].strip()
                        file_type = os.path.splitext(file_tag.text)[1].strip().split(' ')[0]                                        
                        if re.search('download', tag.get_attribute('class')):
                            remove_dup_files(file_name) ## 먼저 지우고
                            download_files.append(re.sub(" ", "", file_name))
                            tag.click() ## 다운로드                        
                            break
            ########################################################################################
            ## 파일 리스트를 첨부할때 웹페이지의 정보를 그대로 가져오면 안됨
            ## 파일명과 실제로 다운로드 파일명이 다른 경우가 있어서 강제로 띄워쓰기를 없애고 비교해야함
            ## 파일 다운로드에 시간이 걸리므로, 완료될 때까지 대기함(제한시간 있음)
            ########################################################################################
            ## 다운로드가 끝날 때까지 대기(제한시간 초과 시 받아진 것만 사용)
            files = wait_for_downloads(download_files)
            
            mail_body = "자세한 내용은 첨부한 HWP 파일을 참고해주세요."
            send_mail(mail_title, mail_body, files)
            sended_mail_dates.append(date)
            sended_mail_subjects.append(subject)            
    return sended_mail_subjects, sended_mail_dates


#############################################################################
## 최종적으로 새로나온 보도자료들을 수집하고 이메일까지 보냄
## 1페이지 이메일 발송 후, 2페이지까지 보도자료가 있으면 페이지 이동도함
#############################################################################
def notice_fsc_press(browser, today, title_keywords=None):
    ## 보도자료가 한번에 10페이지 이상 올라올 가능성은 없음. 그러므로 최대 10페이지까지만 검색함
    last_page_num = 10

    move_to_home(browser, 'fsc_press')
    mail_subject_list, mail_date_list = [], []
    for page_num in range(1, last_page_num+1):  ## page_num = 지금 처리 중인 페이지
        sended_mail_subjects, sended_mail_dates = download_send_mail(browser, today, title_keywords) ## 1페이지는 무조건 작업을 하고
        mail_subject_list.extend(sended_mail_subjects)
        mail_date_list.extend(sended_mail_dates)
        
        ## 다음페이지 넘어갈지 여부 조사
        table_df = get_table_data(browser)
        find_continue = (len(table_df)==len(table_df[table_df.등록일>=today]))        
        if not find_continue:  ## 이 페이지에 기존 자료가 섞여 있으면 여기까지가 끝
            break
        if page_num == last_page_num:  ## 상한까지 다 찼음. 조용히 끝내지 말 것
            print(f'[경고] (금융위 보도자료) 최대 {last_page_num}페이지까지만 확인했습니다. '
                  f'아직 확인하지 않은 자료가 남아 있을 수 있습니다.')
            break
        move_to_page(browser, page_num + 1)
        time.sleep(5) ## 페이지 넘어갈땐 한참 대기
    if len(mail_subject_list)>0:
        sended_mail = pd.DataFrame({'보도자료':mail_subject_list, '날짜':mail_date_list})
        print(f'(금융위 보도자료) 총 {len(sended_mail)}개의 이메일을 발송하였습니다. 상세 리스트는 아래 표를 참고해주세요.')
        display(sended_mail)
    else:
        print('(금융위 보도자료) 발송된 메일이 없습니다.)')