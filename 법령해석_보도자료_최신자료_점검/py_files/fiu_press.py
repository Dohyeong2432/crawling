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

## 해당페이지 표 데이터 수집(제목과 보도일자만)
def get_table_data(browser):
    table_tag = browser.find_element(By.CLASS_NAME, 'bo_list')
    ## 표가 뜰때까지 잠시 대기
    while True: 
        if len(table_tag.find_elements(By.CLASS_NAME, 'bo_li'))>0:
            break
        else:
            time.sleep(0.1)    
    
    dates = []
    subjects = []

    for li_tag in table_tag.find_elements(By.CLASS_NAME, 'bo_li'):
        div_tag = li_tag.find_element(By.CLASS_NAME, 'wrap_cont')

        # 제목은 wrap_cont 안의 첫번째 p 태그    
        subjects.append(div_tag.find_element(By.TAG_NAME, 'p').text.strip())
        ## span 태그 중 date가 클래스명에 들어간 첫번째 태그가 보도일자
        for span_tag in div_tag.find_elements(By.TAG_NAME, 'span'):
            if re.search('date', span_tag.get_attribute('class')):
                dates.append(span_tag.text.strip())
                break        
    table_df = pd.DataFrame({'제목':subjects, '등록일':dates})
    table_df['등록일'] = table_df['등록일'].map(parse_press_date)
    return table_df

## 새로운 보도자료가 나와있는지 확인
def update_check(browser, today):
    move_to_home(browser, 'fiu_press')
    table_df = get_table_data(browser)

    table_df = table_df[table_df.등록일>=today]
    if len(table_df) > 0:        
        print("(금융정보분석원 보도자료) 새로 올라온 자료가 있습니다. 아래 표를 참고해주세요.")
        display(table_df)
    else:
        print("(금융정보분석원 보도자료) 새로운 자료 없습니다.")

## 해당 페이지 번호로 이동
def move_to_page(browser, page_num):
    page_tag = browser.find_element(By.CLASS_NAME, 'paging_g')
    for a_tag in page_tag.find_elements(By.TAG_NAME, 'a'):
        if re.search(r'\d+', a_tag.text):
            get_page_num = "".join(re.findall('[0-9]', a_tag.text))       
        
            if get_page_num==str(page_num):
                a_tag.click()                
                break

## 게시글에 첨부된 파일 중 HWP 파일만 다운로드 하고 리스트 반환
def get_file_list(browser):    
    files = []
    file_tag = browser.find_element(By.CLASS_NAME, 'file_list')    

    loop_count = 0
    while True:
        if loop_count > 50: ## 충분히 기다려도 없으면 파일이 첨부되지 않은 것으로 간주
            break
        if len(file_tag.find_elements(By.TAG_NAME, 'li'))>0:
            break
        else:
            loop_count += 1
            time.sleep(0.1)
    for li_tag in file_tag.find_elements(By.TAG_NAME, 'li'):
        file_name, file_type = os.path.splitext(li_tag.text.strip())        
        remove_dup_files(file_name)
        if re.search('hwp', file_type):
            if re.search('hwpx', file_type) is None:        
                ## a태그 click으로 다운로드가 안됨. 반드시 LINK_TEXT로 접근        
                browser.find_element(By.LINK_TEXT, li_tag.text).click()            
                files.append(re.sub(" ", "", file_name)) ## 파일이름은 공백을 제거하고. 확장자 없이 넘김    
    return files

def download_send_mail(browser, today, title_keywords=None):
    sended_mail_subjects, sended_mail_dates = [], []
    table_tag = browser.find_element(By.CLASS_NAME, 'bo_list')
    ## 표가 뜰때까지 잠시 대기
    while True:         
        if len(table_tag.find_elements(By.CLASS_NAME, 'bo_li'))>0:
            break
        else:
            time.sleep(0.1)    
    
    for li_tag in table_tag.find_elements(By.CLASS_NAME, 'bo_li'):
        div_tag = li_tag.find_element(By.CLASS_NAME, 'wrap_cont')
        subject_tag = div_tag.find_element(By.TAG_NAME, 'p')
        subject = subject_tag.text.strip()
        
        ## span 태그 중 date가 클래스명에 들어간 첫번째 태그가 보도일자
        for span_tag in div_tag.find_elements(By.TAG_NAME, 'span'):
            if re.search('date', span_tag.get_attribute('class')):
                date = parse_press_date(span_tag.text)
                break        
        if date>=today and subject_matches(subject, title_keywords):
            ## 자바스크립트 함수 호출을 위해, 함수평 가져오고 실행
            java_link = subject_tag.find_element(By.TAG_NAME, 'a').get_attribute('href')
            browser.execute_script(java_link)
    
            ## 파일 다운로드 시간이 걸릴수 있으므로 좀 대기할 것
            download_files = get_file_list(browser)  ## 게시글의 파일 다운로드 및 리스트 가져오기                
            ## 다운로드가 끝날 때까지 대기(제한시간 초과 시 받아진 것만 사용)
            files = wait_for_downloads(download_files)
            mail_title = f'(금융정보분석원 보도자료)_{date}_{subject}'
            mail_body = '자세한 내용은 첨부한 HWP 파일을 참고해주세요.'
            send_mail(mail_title, mail_body, files)
            sended_mail_subjects.append(subject)
            sended_mail_dates.append(date)
            time.sleep(0.1)
            browser.back()
    return sended_mail_subjects, sended_mail_dates


#############################################################################
## 최종적으로 새로나온 보도자료들을 수집하고 이메일까지 보냄
## 1페이지 이메일 발송 후, 2페이지까지 보도자료가 있으면 페이지 이동도함
#############################################################################
def notice_fiu_press(browser, today, title_keywords=None):
    ## 보도자료가 한번에 10페이지 이상 올라올 가능성은 없음. 그러므로 최대 10페이지까지만 검색함
    last_page_num = 10    
    
    move_to_home(browser, 'fiu_press')
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
            print(f'[경고] (금융정보분석원 보도자료) 최대 {last_page_num}페이지까지만 확인했습니다. '
                  f'아직 확인하지 않은 자료가 남아 있을 수 있습니다.')
            break
        move_to_page(browser, page_num + 1)
        time.sleep(5) ## 페이지 넘어갈땐 한참 대기
    if len(mail_subject_list)>0:
        sended_mail = pd.DataFrame({'보도자료':mail_subject_list, '날짜':mail_date_list})
        print(f'(금융정보분석원 보도자료) 총 {len(sended_mail)}개의 이메일을 발송하였습니다. 상세 리스트는 아래 표를 참고해주세요.')
        display(sended_mail)
    else:
        print('(금융정보분석원 보도자료) 발송된 메일이 없습니다.')