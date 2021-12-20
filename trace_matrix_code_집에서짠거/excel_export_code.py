import pandas as pd
import openpyxl
from openpyxl import Workbook
import subprocess
import math
import os
import re
import chardet
from enum import Enum
import datetime
import time
import csv

start = time.time()


DocID = '1226890'     # 1226890
SysRS1ID = '1351133'
SysRS2ID = '1356188'
SysRS3ID = '1430240'
SwRS1ID = '1392803'
SwRS2ID = '1394006'
SwRS3ID = '1469578'
SysTCID = '1454824'
SwTCID = '1464826'

now = datetime.datetime.now()
today = now.strftime('%Y-%m-%d %H:%M')
nowtime = now.strftime('%m%d_%H%M')
outputFileName_GetTestResult = 'getTestResult_Test'+nowtime+'.xls'
outputFileName_GetTestResult_SysSwTS = 'getTestResult_Test_SysSwTS'+nowtime+'.xls'

def main():

    global row_cr
    global row_tc
    global user_id_pw
    global implementation_criteria
    global minor_criteria
    global skip_string
    global LukID_temp

    implementation_criteria = 12
    minor_criteria = 'Patch#2'

    user_id_pw = [' ', ' ']
    # Get test session IDs
    test_session_txt = "testsession_v104.5.txt"
    user_name_txt = "user.txt"
    read_xlsx = "read.xlsx"
    DocID_Info_xls = "DocID_Info.xls"
    SysID_Info_xls = "SysID_Info.xls"
    SwID_Info_xls = "SwID_Info.xls"
    SysSwTS_Info_xls = "SysSwTSID_Info.xls"
    TestResult_csv = "TestResult_Info.csv"

    #QueryDefinition_GetTestResult = '((field["Document ID"]='+DocID+')and(field["Project"]="/Schaeffler MCA LCU")and(item.live)and(item.meaningful)and("disabled not"(field["Category"]="Heading","Comment")))'
    #QueryDefinition_GetTestResult_SysSwTS = '((field["Document ID"]='+SwTCID+')and(field["Project"]="/Schaeffler MCA LCU")and(item.live)and(item.meaningful)and("disabled not"(field["Category"]="Heading","Comment")))'

    # 추출할 item filed 정의
    #itemExportFields = '"Document ID",ID,"A15 LuK ID",Text,"A05 Safety Integrity","A25 Status Commitment Supplier - MCA LG","A25 Status Commitment Supplier - MCA LG HMC","A27 Delivery Date","Short Description","Decomposes To"'
    #itemExportFields_SysSwTS = '"Document ID",ID,"ENG ID","Test Method"'

    # exporting Non traced items of specific document
    #export_doc_cmd = 'im exportissues --outputFile=' + outputFileName_GetTestResult + ' --fields=' + itemExportFields + ' --sortField=Type --queryDefinition=' + QueryDefinition_GetTestResult
    #export_doc_cmd_SysSwTS = 'im exportissues --outputFile=' + outputFileName_GetTestResult_SysSwTS + ' --fields=' + itemExportFields_SysSwTS + ' --sortField=Type --queryDefinition=' + QueryDefinition_GetTestResult_SysSwTS
    #subprocess.Popen(export_doc_cmd)
    #subprocess.Popen(export_doc_cmd_SysSwTS)

    try:
        with open(test_session_txt, 'rt') as in_file:
            for line in in_file:
                test_session_id = line  # test_session_txt 파일에 1줄마다 읽어드림
                test_session_id = test_session_id.split(',')  # ,로 문자열 리스트화에서 구별
            len_test_session_id = len(test_session_id)
    except:
        print("Please make a testsession.txt file with test session id. e.g. 1495334, 1495339, 1495343, 1495344")  # test_seesion_txt 파일에 아무것도 데이터가 없을때 또는 ,로 구분안갈때
        subprocess.check_output("pause", shell=True)
        return

    f = open(TestResult_csv, 'w', encoding='EUC-KR', newline='')
    wr = csv.writer(f)
    wr.writerow(["Session ID", "TC ID", "Test Result"])

    itemExportFields_TestResult = 'sessionID,caseID,verdict'

    for j in range(0, len_test_session_id):
        result_test_cmd = 'tm results' + ' --sessionID=' + test_session_id[j].strip() +' --fields=' + itemExportFields_TestResult

        result = subprocess.check_output(result_test_cmd)
        result = result.splitlines()

        for line in result:  # Store each line in a string variable "line"
            # parse id decomposed to, satisfied by, validated by
            try:
                encode_type = chardet.detect(line)
                if encode_type['encoding'] is not None:
                    line = line.decode(encode_type['encoding'])             # encode_type['encoding'] = EUC-KR
                    line = line.split('\t', maxsplit=3)
                    #ssprint(line)
                    wr.writerow(line)
                else:
                    line = line.decode('EUC-KR')
            except:
                print("problem is occured. id", id)

    f.close()

main()

end = time.time()
print(f"TestResult파일이 만드는데 까지 걸리는 시간 : {end - start:.5f} sec")