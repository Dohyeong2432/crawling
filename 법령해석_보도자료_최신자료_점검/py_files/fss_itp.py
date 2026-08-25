from win32com.client import constants
import pandas as pd
import numpy as np
import locale, glob, math, time, random, warnings, datetime, requests, ssl, smtplib
import requests, json, time, re, os
import xlwings as xw

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.encoders import encode_base64

from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup as bs

from py_files.common_functions import *
os.environ['WDM_SSL_VERIFY'] = '0' ## ChromeDriverManager().install() 실행 시 verify = False 옵션
warnings.filterwarnings('ignore')

## 환경변수 설정 값은 common_functions.py 에서 한 번만 읽어와 공유함
## (from py_files.common_functions import * 로 SEND_ADDR/DOWNLOAD_FOLDER 등이 들어옴)

def get_num_tag(browser):
    num_tag = False
    for tag in browser.find_elements(By.TAG_NAME, 'div'):
        if tag.get_attribute('class')=="dataTables_paginate paging_full_numbers":
            num_tag = tag
    return num_tag if num_tag else None

def move_to_first_page(browser):
    move_to_home(browser, 'fss_itp')

    num_tag = get_num_tag(browser)   
    
    for a_tag in num_tag.find_elements(By.TAG_NAME, 'a'):        
        if re.search('first', a_tag.get_property('id')):
            a_tag.click()
    time.sleep(0.3)

## 법령 해석 리스트의 가장 마지막 페이지 번호 호출
def get_last_page_num(browser):
    move_to_home(browser, 'fss_itp')
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
def get_post_info(browser):
    col_name = browser.find_elements(By.TAG_NAME, 'table')[0].find_element(By.TAG_NAME, 'td').text ## 제목은 첫번째 표에서 가져옴(두번째 표에 제목없는 경우가 있음)
    table_tag = browser.find_elements(By.TAG_NAME, 'table')[1] ## 세부 페이지에서 가져올 표는 두번째 표(회신)임
    tbody_tag = table_tag.find_element(By.TAG_NAME, 'tbody')    
    page_df = pd.DataFrame(columns=[col_name])
    
    for tr in tbody_tag.find_elements(By.TAG_NAME, 'tr')[1:]:    
        idx_name = tr.find_element(By.TAG_NAME, 'th').text.strip()
        td_text = tr.find_element(By.TAG_NAME, 'td').text.strip()
        if idx_name=='첨부파일': ## 첨부파일은 다운로드 받음
            remove_dup_files(os.path.splitext(td_text)[0])
            tr.find_element(By.TAG_NAME, 'td').find_element(By.TAG_NAME, 'a').click()
            
        page_df.loc[idx_name, col_name] = td_text    
    update_date = pd.to_datetime(page_df.loc['회신일'].values[0]).strftime("%Y%m%d")
    law_name = page_df.columns[0]
    law_name = re.sub(r'[\\/*?:"<>|]', '', law_name)

    ## 정리한 page_df를 파일로 저장하고 편집
    page_info_file = f"./update_list/{update_date}_{law_name}.xlsx"
    page_df.to_excel(page_info_file)
    app = xw.App(visible=True)
    wb = app.books.open(page_info_file)
    sht = wb.sheets[0]        
    # A/B열 조정
    sht.range("B:B").api.WrapText = True
    sht.range("B:B").api.VerticalAlignment = constants.xlTop
    sht.range("B:B").column_width = 150
    sht.range("B:B").api.EntireRow.AutoFit()
    wb.save()
    wb.close()
    app.quit()

    mail_title = f"(법령해석포털)_{update_date}_{law_name}"
    mail_body = "본문의 내용은 첨부한 엑셀파일을 참고해주세요."     

    files = [page_info_file] + list(os.path.join(DOWNLOAD_FOLDER, file) for file in page_df.loc['첨부파일'].values.tolist())    
    return mail_title, mail_body, files


## 각 항목들을 클릭하고 페이지 정보를 수집한 후 메일까지 발송
def get_page_laws_info(browser, table_df):
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
                        mail_title, mail_body, files = get_post_info(browser) ## 상세 페이지의                         
                        time.sleep(0.2)
                        browser.back()                        
                        send_mail(mail_title, mail_body, files)
                        break        
    
        time.sleep(0.2)
    

## 업데이트된 법령이 있는지만 체크
def update_check(browser, today):
    last_page_num = get_last_page_num(browser)
    table_df, tbody_tag = get_table_data(browser)
    table_df['등록일'] = table_df['등록일'].map(lambda x: pd.to_datetime(x).strftime("%y%m%d"))
    
    table_df = table_df[table_df.등록일>=today]
    if len(table_df)>0:
        print("(법령해석) 새로 올라온 정보가 있습니다. 아래 표를 참고해주세요.")
        display(table_df)        
    else:
        print('(법령해석) 새로 올라온 정보는 없습니다.')

####################################################################################
## 업데이트된 법령이 있는지 체크하고 있으면 최종적으로 실행해 수집후 이메일 보내는 함수
def notice_fss_ipt(browser, today):
    move_to_first_page(browser)
    last_page_num = get_last_page_num(browser)

    table_list = []
    sended_mails = 0
    ## 모든페이지를 돌면서 업데이트된 정보가 있는지 파악함
    for page_num in range(1, last_page_num+1):
        if page_num!=1: ## 첫페이지는 무조건 테이블을 수집
            num_tag = get_num_tag(browser) ## 화면 하단 페이지 태그 부분을 가져와서 해당 페이지로 이동
            for a_tag in num_tag.find_elements(By.TAG_NAME, 'a'):        
                find_page_num = a_tag.text.strip()
                if find_page_num:
                    if page_num==int(find_page_num):
                        a_tag.click()
                        break
        while True:
            try:
                table_df, tbody_tag = get_table_data(browser)
                break
            except:
                time.sleep(0.05)            
        table_df['등록일'] = table_df['등록일'].map(lambda x: pd.to_datetime(x).strftime("%y%m%d"))

        ## 전체 테이블과 가져올 테이블의 수가 같으면 다음페이지도 조사
        ## 왜냐하면 업데이트된 항목이 1페이지 이상일 수도 있으므로
        find_continue = (len(table_df)==len(table_df[table_df.등록일>=today])) 
        table_df = table_df[table_df.등록일>=today]
        if len(table_df) > 0: ## 찾아야할 테이블이 있다면
            ## 각 항목들을 클릭해서 파일을 다운로드 하고, 본문을 수집한 후 메일까지 발송
            get_page_laws_info(browser, table_df)
            sended_mails = sended_mails + len(table_df)
            table_list.append(table_df)
        if find_continue:
            time.sleep(5) ## 페이지 넘어갈땐 충분히 대기
            continue            
        else:        
            break
    if sended_mails>0:
        print(f'(법령해석) {sended_mails} 개의 메일을 발송하였습니다. 상세 리스트는 아래 표를 참고해주세요.')
        display(pd.concat(table_list).reset_index(drop=True))
    else: 
        print('(법령해석) 발송된 메일이 없습니다.')