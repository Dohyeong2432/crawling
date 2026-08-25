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

## 해당페이지 표 데이터 수집
def get_table_data(browser):
    table_tag = browser.find_element(By.TAG_NAME, 'table')
    tbody_tag = table_tag.find_element(By.TAG_NAME, 'tbody')
    
    table_cols = []
    for th_tag in table_tag.find_elements(By.TAG_NAME, 'th'):
        if th_tag.text:
            table_cols.append(th_tag.text.strip())
    
    table_df = pd.DataFrame(columns=table_cols)    
    
    for row_idx, tr_tag in enumerate(tbody_tag.find_elements(By.TAG_NAME, 'tr')):
        for col_idx, td_tag in enumerate(tr_tag.find_elements(By.TAG_NAME, 'td')):
            table_df.loc[row_idx, table_cols[col_idx]] = td_tag.text.strip()

    table_df['등록일'] = table_df['등록일'].map(lambda x: pd.to_datetime(x).strftime("%y%m%d"))

    return table_df

## 새로운 보도자료가 나와있는지 확인
def update_check(browser, today):
    move_to_home(browser, 'fss_press')
    table_df = get_table_data(browser)
    table_df = table_df[table_df.등록일>=today]
    if len(table_df) > 0:        
        print("(금감원 보도자료)새로 올라온 자료가 있습니다. 아래 표를 참고해주세요.")
        display(table_df)
    else:
        print("(금감원 보도자료) 새로 올라온 자료는 없습니다.")

## 해당 페이지 번호로 이동
def move_to_page(browser, page_num):
    ## 하단 페이지 이동 영역을 찾은 후
    for ul_tag in browser.find_elements(By.TAG_NAME, 'ul'):
        class_name = ul_tag.get_attribute('class')
        if class_name:
            if re.search('pagination-centered', class_name):
                break    
    # 페이지 이동
    for li_tag in ul_tag.find_elements(By.TAG_NAME, 'li'):    
        if li_tag.text==str(page_num):
            li_tag.find_element(By.TAG_NAME, 'a').click()
            break

## 게시판 본문에서 본문의 내용과 첨부할 파일리스트 전달
def get_post_info(browser):
    while True: ## 게시글이 정상적으로 올라올때까지 대기
        if browser.find_elements(By.CLASS_NAME, 'dbdata'):
            break
        else:
            time.sleep(0.1)
            continue
    
    body_tag = browser.find_element(By.CLASS_NAME, 'dbdata')
    mail_body = body_tag.text
    
    file_tags = []
    for div_tag in browser.find_elements(By.CLASS_NAME, 'file-list__set__item'):
        file_tags.append(div_tag)
    
    download_files = []
    for file_tag in file_tags:
        file_name, file_type = os.path.splitext(file_tag.text)    
        remove_dup_files(file_name) ### 기존에 받은게 있으면 지우고
        if re.search('hwp', file_type):    ## 모든 HWP 파일을 다 다운로드 받음
            file_tag.find_element(By.TAG_NAME, 'a').click()  ## 첫번째 a태그만 다운로드
            download_files.append(re.sub(" ", "", file_name))
    
    ## 파일 다운로드에 시간이 걸리므로, 완료될 때까지 대기함(제한시간 있음)
    ## 다운로드가 끝날 때까지 대기(제한시간 초과 시 받아진 것만 사용)
    files = wait_for_downloads(download_files)
    browser.back() ## 뒤로 가고 종로
    return mail_body, files ## 메일 본문의 글과, 첨부할 파일 리스트

#############################################################################
## 최종적으로 새로나온 보도자료들을 수집하고 이메일까지 보냄
## 1페이지 이메일 발송 후, 2페이지까지 보도자료가 있으면 페이지 이동도함
#############################################################################
def notice_fsc_press(browser, today):
    ## 보도자료가 한번에 10페이지 이상 올라올 가능성은 없음. 그러므로 최대 10페이지까지만 검색함
    last_page_num = 10

    move_to_home(browser, 'fss_press')
    mail_subject_list, mail_date_list = [], []

    for page_num in range(1, last_page_num+1):  ## page_num = 지금 처리 중인 페이지
        ## 칼럼 정보 취득
        table_df = get_table_data(browser)
        table_cols = table_df.columns.tolist()
        
        for tag_idx in range(len(table_df)): ## table_df의 행 갯수가 현재 페이지 게시글의 수임
            ## 본문에 들어갔다 돌아오면 모든 태그가 사라지므로 새로 세팅
            table_tag = browser.find_element(By.TAG_NAME, 'table')
            tbody_tag = table_tag.find_element(By.TAG_NAME, 'tbody')
            tr_tag = tbody_tag.find_elements(By.TAG_NAME, 'tr')[tag_idx]
        
            td_tag_list = tr_tag.find_elements(By.TAG_NAME, 'td')
        
            date = pd.to_datetime(td_tag_list[table_cols.index('등록일')].text).strftime("%y%m%d")
            subject = td_tag_list[table_cols.index('제목')].text
            mail_title = f"(금감원보도자료)_{date}_{subject}" ## 메일 제목 양식 수정
            
            if date>=today: ## 오늘 이후로 업데이트 된 내용이면 첨부
                td_tag_list[table_cols.index('제목')].click() ## 해당 본문 클릭
                mail_body, files = get_post_info(browser) ## 본문의 내용과 파일을 다운로드 한후
                send_mail(mail_title, mail_body, files) ## 이메일 발송

                ## 발송내역 체크를 위해 별도 저장
                mail_subject_list.append(subject)
                mail_date_list.append(date)
                time.sleep(0.5)    

        find_continue = (len(table_df)==len(table_df[table_df.등록일>=today]))        
        if not find_continue:  ## 이 페이지에 기존 자료가 섞여 있으면 여기까지가 끝
            break
        if page_num == last_page_num:  ## 상한까지 다 찼음. 조용히 끝내지 말 것
            print(f'[경고] (금감원 보도자료) 최대 {last_page_num}페이지까지만 확인했습니다. '
                  f'아직 확인하지 않은 자료가 남아 있을 수 있습니다.')
            break
        move_to_page(browser, page_num + 1)
        time.sleep(5) ## 페이지 넘어갈땐 한참 대기

    if len(mail_subject_list)>0:
        sended_mail = pd.DataFrame({'보도자료':mail_subject_list, '날짜':mail_date_list})
        print(f'총 {len(sended_mail)}개의 이메일을 발송하였습니다. 상세 리스트는 아래 표를 참고해주세요.')
        display(sended_mail)
    else:
        print("(금감원 보도자료) 발송된 메일이 없습니다.")