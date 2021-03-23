#-*-coding:utf-8
# ------------------------------------------
# 파이썬 3.7 버전 기준으로 작성됨
# 특정 조건의 파일 리스트를 읽어와 다운로드함
# ------------------------------------------
# 최초 작성 : 2020.04 이철희
# ------------------------------------------
# 2020.11 이철희 고용보험신고를 위한 게약서 갑지와 납부확인서 다운로드
# ------------------------------------------
# pyinstaller --ico=gs.ico --onefile kcom_filedn.py
# ep_user_pin = '' #EP 로그인 PIN번호
# MS-SQL 서버
mssql_server_ip = '165.243.39.140'
mssql_server_pw = 'dptmzbdpf!'

# 대기시간 조절을 위한 셀렌 부속 모듈
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

import time
import sys
from selenium import webdriver
import getpass

# 계약 파일 저장을 위한 리퀘스트 등
from urllib.request import urlretrieve
import re
import pdfkit # PDF변환 모듈
file_folder = 'C:/PMS/KCOM/' # 미리 만들어 놓을것

# DB접속을 위한 오라클/MS-SQL
import os
#import cx_Oracle
import pymssql

from PyPDF2 import PdfFileMerger # PDF 합치기

config = pdfkit.configuration(wkhtmltopdf='C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe') #사전에 인스톨 필요

SHIFT=1

han=re.compile('[^ㄱ-ㅣ가-힣]+') #한글
# -------------------------------------
# 한글 유니코드 인코딩
def han_encode( raw_str ):
    enc_str = ''
    hans=han.findall(raw_str)
    for i in raw_str:
        if ( i in str(hans) ):
            enc_str += i
        else:
            enc_str += str(hex(ord(i))).replace('0x','%u') # 한글만 인코딩함
    return enc_str

def decrypt( raw ):
    ret = ''
    for char in raw:
        ret+=chr( ord(char)-SHIFT )
    return ret

# -------------------------------------
# DB 조회시 컬럼을 사용하기 위한 정의
def makeDictFactory(cursor):
   columnNames = [d[0] for d in cursor.description]

   def createRow(*args):
      return dict(zip(columnNames, args))
 
   return createRow

# -------------------------------------
# 계약첨부파일 다운로드 
def file_down( row, file_url ):
    
    try:
        file_name = file_url.split('=')[-1] # 파일 이름만 가져옴
        new_file_name = file_name
    except Exception as ex:
        print('계약파일오류: '+row['HJCD'] + "-" + row['UPGJ'] + "-" + row['GJSEQ']+str(ex))
        return 'none'

    url = 'http://pms.gsconst.co.kr/PMS.WEB/PMS.WEB.HY/PMV42010_DOWN.aspx' + han_encode( file_url ).replace(' ','+') #+ '&SYS=KSC'
    # url = 'http://pms.gstpms.com/eps_files/' + han_encode( file_url ).replace(' ','+')
    
    try:
        urlretrieve( url, file_folder + file_name )
        if ( '.htm' in file_name ):
            # htm -> pdf 변환
            new_file_name = file_name.replace('.htm','.pdf')
            # pdfkit 출력을 devnull 로 잠시 보내기...(pdfkit 실행시 stdout mute 시킴)
            old_stdout = sys.stdout
            sys.stdout = open(os.devnull, "w")
            pdfkit.from_file( file_folder + file_name, file_folder + new_file_name ,configuration=config)
            sys.stdout = old_stdout
            # pdf 변환이 성공했으면 원래 파일 삭제
            os.remove( file_folder + file_name )

        #파일이름 앞에 현장-공종-순차코드 세팅
        os.rename( file_folder + new_file_name,  file_folder + row['HJCD'] + "-" + row['UPGJ'] + "-" + row['GJSEQ'] + "_" + new_file_name )    
    except Exception as ex:
        print('---- 파일다운로드 오류 ---- ' + url + ':' + str(ex) )
 
    return new_file_name

# -------------------------------------
# 다운로드한 파일중 PDF 변환된 파일을 하나로 합침
def merge_pdf( row, new_file_name ):
    
    merger = PdfFileMerger()


    try:
        # 파일앞에 이름이 현장+공종+순차인 pdf를 하나로 합침(ex:현장공종순차_고용보험신고용서류)
        if os.path.exists(file_folder):
            for file in os.scandir(file_folder):
                if ( file.is_dir() ): # 하위 폴더는 유지하고 파일만 삭제
                    pass
                else:
                    if ( file.name.startswith(row['HJCD'] + "-" + row['UPGJ'] + "-" + row['GJSEQ']) ):
                        merger.append(file_folder + file.name)
                    else:
                        #병합 파일 이름과 동일한 이름이 존재하는 경우 해당 파일을 먼저 지운다.
                        if ( file.name.startswith('GSENC-' + row['HJNM'] +"-"+ row['HJCD'] + "-" + row['UPGJ'] + "-" + row['GJSEQ']) ):
                            os.remove(file.path)

            merger.write(file_folder + 'GSENC-' + row['HJNM'] +"-"+ row['HJCD'] + "-" + row['UPGJ'] + "-" + row['GJSEQ'] 
                         + "-" + row['CONM'] + "_" + new_file_name + '.pdf' )
            merger.close()
    except Exception as ex:
        print('-----PDF합치기오류-----' + ':' + str(ex) )
        
    return 0

# -------------------------------------
# 파일 삭제(특정 폴더의 폴더 제외한 모든 파일 삭제)
def removeAllFile(path,delee):
    if os.path.exists(path):
        for file in os.scandir(path):
            if ( file.is_dir() ): # 하위 폴더는 유지하고 파일만 삭제
                pass
            else:
                if ( delee=="TEMP" and file.name.startswith('GSENC') ):
                    pass
                else:
                    os.remove(file.path)

        return True
    else:
        return False

print( time.ctime() )
print( "------------고용보험 협력사 신고 파일(계약서,납부확인서) 다운로드 자동화 프로그램 시작-----------" )
print( "현재일 기준 계약일 2개월 이내의 고용보험 협력사납부 계약건(계약완료기준)에 대해" )
print( "신청번호가 없거나 신청번호는 있어도 승인일이 없는 건의 계약서갑지와 납부확인서를" )
print( "다운받아 GSENC 현장명+현장코드+공종코드+순차+협력사명으로 시작하는 pdf 파일로 합쳐 저장합니다." )
print( "-----------" )

# # 로그인을 위한 PIN번호 입력
#ep_user_pin = input("Input EP PIN Number :")
ep_user_pin = getpass.getpass("Input EP PIN Number :")

# # GS건설 EP 로그인
browser = webdriver.Ie('C:/PMS/IEDriverServer')
# #browser = webdriver.Chrome('C:/PMS/chromedriver')
browser.set_window_position(0, 0)
browser.set_window_size( 1280, 1024 )

browser.get('http://ep.gsenc.com')    

element = browser.find_element_by_id('apn_ori_pw')
browser.execute_script("arguments[0].value='" + ep_user_pin + "'",element)
browser.find_element_by_class_name('btn_login_certi').click()

# time.sleep(5)

# 인증성공여부 확인
try:
    WebDriverWait(browser, 10).until( EC.presence_of_element_located( (By.ID,'wrapper') ) )
    browser.find_element_by_id('wrapper')
    print('EP 인증 성공 ')
except Exception as ex:
    print('EP 인증 실패 : '+ str(ex))
    sys.exit()

time.sleep(5)

# 다운로드폴더 탐색창 띄우기
if not os.path.isdir(file_folder):
    os.mkdir(file_folder)
    
os.startfile(os.path.realpath(file_folder))

# 기존파일 삭제
removeAllFile(file_folder,'ALL')

# DB 연결
os.putenv('NLS_LANG','KOREAN_KOREA.KO16MSWIN949')
# try:
#     cx_Oracle.init_oracle_client(lib_dir="C:\PMS\instantclient_19_8")
# except Exception as ex:
#     print('Oracle Connection :' + str(ex))

# DB 연결(MS-SQL)    
try:
    connection=pymssql.connect(server=mssql_server_ip, user='seldb', 
                               password=mssql_server_pw, database='SELDB', as_dict=True, charset='EUC-KR')
except Exception as ex:
    print('MS-SQL Connection :' + str(ex))
    
# connection = cx_Oracle.connect("pmsplus",decrypt("yvgj8:wmi"),"165.243.39.155:1531/dbsvr08")

#-------------------------------------------------------------
# 파일 리스트 조회
# PMVT420 현설기준 GUBUN2 = 'HS' 파일SEQ 2 = 현장견적조건
#-------------------------------------------------------------
main_cursor = connection.cursor()
#main_cursor.execute("set name utf8")
main_cursor.execute("""
                    SELECT * FROM OPENQUERY(DBSVR08,'
SELECT A.HJCD, A.UPGJ, A.GJSEQ, A.COCD,
        A.UPGJ||chr(45)||A.GJSEQ||chr(45)||A.COCD AS CON_CODE,  -- 계약코드
        A.KYDATE                          AS ORG_KYDATE, -- 최초계약일
        DECODE(NVL(A.KYAMT1,0),0,A.KYAMT,A.KYAMT1) AS KYAMT,
        REPLACE(FN_CT000_GET_HJNM(A.HJCD),chr(47),chr(32)) AS HJNM,
        FN_CT090_GET_CONM(A.COCD) AS CONM,
        FN_CT100_GET_UPNM(A.UPGJ) AS UPNM, 
        NVL(B.GY_NUM,chr(32)) AS GY_NUM, -- 고용보험관리번호
        B.CON_STATUS, C.GUBUN, 
        NVL(C.GYGJ_FG,chr(78)) AS GYGJ_FG, -- 신고대상여부
        C.REQ_DATE,  -- 신청일자
        C.REQ_NUM,   -- 신청번호
        C.LIC_TYPE, C.LIC_NUM, C.LIC_DATE, -- 면허종류,번호,등록일
        B.HDDATE, -- 하도급통보일
        ( SELECT ''?file_path=upload/''||FILE_PATH||chr(38)||''file_name=''||FILE_NM 
           FROM PMVT420 
          WHERE HJCD = A.HJCD AND UPGJ = A.UPGJ AND GJSEQ = A.GJSEQ AND COCD = A.COCD
            AND HCHNO = B.HCHNO AND GUBUN LIKE ''A%'' AND ROWNUM = 1 )       FILE_KY, -- 계약서갑지
        ( SELECT ''?file_path=upload/''||FILE_PATH||chr(38)||''file_name=''||FILE_NM 
           FROM PMVT420 
          WHERE HJCD = A.HJCD AND UPGJ = A.UPGJ AND GJSEQ = A.GJSEQ AND COCD = A.COCD
            AND HCHNO = B.HCHNO AND GUBUN LIKE ''N%'' AND ROWNUM = 1 )      FILE_RE, -- 납부확인서
        1
   FROM PMAT000 A, PMAT010 B, PMAT060 C  
  WHERE A.HJCD = B.HJCD  
    AND A.UPGJ = B.UPGJ  
    AND A.GJSEQ = B.GJSEQ
    AND A.COCD = B.COCD  
    AND A.HJCD = C.HJCD       
    AND A.UPGJ = C.UPGJ     
    AND A.GJSEQ = C.GJSEQ 
    AND A.COCD = C.COCD 
    AND A.HJCD NOT LIKE ''PMS%''
    AND A.HJCD NOT LIKE ''XI%''
    AND B.CON_STATUS IN ( ''H'',''E'' )
    -- 계약일 2개월 이내의 건만 확인함
    AND B.KYDATE >= TO_CHAR(ADD_MONTHS(SYSDATE,-2),''YYYYMMDD'')
    AND ( TRIM(C.REQ_NUM) IS NULL -- 신청번호가 없거나 (신청대상)
       -- 신청번호는 있으나 승인일이 없는건(승인번호 확인 대상)
       OR TRIM(C.REQ_DATE) IS NOT NULL AND TRIM(C.DATE1) IS NULL )
    AND B.HCHNO = ''00'' 
    AND (C.UPGJ,C.GJSEQ,C.COCD,C.CDATE) IN ( 
         SELECT UPGJ,GJSEQ,COCD,MAX(CDATE) 
           FROM PMAT060            
          WHERE HJCD = A.HJCD   
            AND GUBUN IN (''2'')  -- 협력사납부만 
          GROUP BY UPGJ,GJSEQ,COCD )        
  ORDER BY B.UPGJ, B.GJSEQ, B.COCD, B.HCHNO  
'
)
    """
    )

# main_cursor.rowfactory = makeDictFactory(main_cursor) # oracle
# ms-sql은 코넥트할때...

for row in main_cursor:
    if ( len(str(row['FILE_KY'])) > 10 ):
        file_name = file_down( row, row['FILE_KY'] )
    if ( len(str(row['FILE_RE'])) > 10 ):
        file_name = file_down( row, row['FILE_RE'] )
    # PDF로 변환된 파일을 하나로 합친다
    merge_pdf( row, '고용보험신고용' )
    print('...processing...')
        
# 임시파일 삭제(GSENC로 시작하지 않는 파일들)
removeAllFile(file_folder,'TEMP')

print('files downloaded. see ' + file_folder )