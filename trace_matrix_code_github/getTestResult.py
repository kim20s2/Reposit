import subprocess
import re
import chardet
import datetime
import time
#import logging
#import logging.handlers

now = datetime.datetime.now()
today = now.strftime('%Y-%m-%d %H:%M')
nowtime = now.strftime('%m%d_%H%M')
outputFileName_GetTestResult = 'getTestResult_Test'+nowtime+'.xls'

class _RegExLib:
    """Set up regular expressions"""
    # use https://regexper.com to visualise these if required
    #_reg_engid = re.compile('ENG ID: (.*)\n')
    #_reg_decomposes = re.compile('Decomposes To: (.*)\n')
    #_reg_luk = re.compile('A15 LuK ID: (.*)\n')
    #_reg_validates = re.compile('Validated By: (.*)\n')
    #_reg_satisfies = re.compile('Satisfied By: (.*)\n')
    #_reg_url = re.compile('URL: (.*)\n')
    #_reg_text = re.compile('Text: (.*)\n')
    _reg_verdict = re.compile('Verdict: (.*)')
    def __init__(self, line):
        # check whether line has a positive match with all of the regular expressions
        # line = line.decode("utf-8", "ignore")
        #self.engid = self._reg_engid.match(line)
        #self.decomposes = self._reg_decomposes.match(line)
        #self.luk = self._reg_luk.match(line)
        #self.validates = self._reg_validates.match(line)
        #self.satisfies = self._reg_satisfies.match(line)
        #self.url = self._reg_url.match(line)
        #self.text = self._reg_text.match(line)
        self.verdict = self._reg_verdict.match(line)

def GetTestResultOneSession(SessionID, TestCaseID):
    # remove white space and ,
    SessionID = SessionID.strip()
    TestCaseID = TestCaseID.strip()
    TestCaseID = TestCaseID.replace('?', '')
    SessionID = str(SessionID)
    TestCaseID = str(TestCaseID)
    # view_test_result_command = 'tm viewresult --user=' + str(user_id_pw[0]) + ' --password=' + str(user_id_pw[1]) + ' ' '--sessionID='+SessionID+' '+TestCaseID       # 로그인 사용
    view_test_result_command = 'tm viewresult --sessionID=' + SessionID + ' ' + TestCaseID                                                                                    # 로그인 사용 X
    try:
        result_cmd = subprocess.check_output(view_test_result_command, shell=True, stderr=subprocess.STDOUT)
        #result_cmd2 = subprocess.run(view_test_result_command)
    except:
        return 'NT'
    result_cmd = result_cmd.splitlines()
    result = 'NT'

    QueryDefinition_GetTestResult = '((field["ID"]='+SessionID+')and(field["Project"]="/Schaeffler MCA LCU")and(item.live)and(item.meaningful)and("disabled not"(field["Category"]="Heading","Comment")))'

    # 추출할 item filed 정의
    itemExportFields = SessionID
    # exporting Non traced items of specific document
    #export_doc_cmd = 'im exportissues --outputFile=' + outputFileName_MCA_Full + ' --fields=' + itemExportFields + ' --sortField=Type --queryDefinition=' + QueryDefinition_notTrace
    export_doc_cmd = 'im exportissues --outputFile=' + outputFileName_GetTestResult + ' --fields=State' + ' --sessionID=' + itemExportFields + ' --sortField=State --queryDefinition=' + QueryDefinition_GetTestResult
    subprocess.Popen(export_doc_cmd)

    for line in result_cmd:  # Store each line in a string variable "line"
        # parse id decomposed to, satisfied by, validated by
        try:
            encode_type = chardet.detect(line)
            if encode_type['encoding'] is not None:
                line = line.decode(encode_type['encoding'])
            else:
                line = line.decode('EUC-KR')
            reg_match = _RegExLib(line)
        except:
            print("problem is occured. id", id)
            print(line)
        if reg_match.verdict:
            # write to excel to CR ID column
            result = reg_match.verdict.group(1)

    return str(result)

def GetTestResultSessions(SessionID, TestCaseID):
    result_session_id = 0
    for single_session_id in SessionID:
        # logging_tr_seaching_text = str(datetime.datetime.now()) + ' searching Case ' + str(TestCaseID) + ' in Session ' + str(single_session_id)
        # logging.info(logging_tr_seaching_text)
        # logging.StreamHandler(logging_tr_seaching_text)
        # print(logging_tr_seaching_text, end="\r")
        # print(logging_tr_seaching_text)
        result = GetTestResultOneSession(single_session_id, TestCaseID)
        result_session_id = single_session_id
        if result == 'Passed':
            break
        elif result == 'Failed':
            break
    if result == 'NT':
        return result
    return result + '(session : ' + str(result_session_id) +')'
    

# def GetTestResultSessions(user_id_pw, SessionID, TestCaseID):
#     tested_session_id = 0
#     for single_session_id in SessionID:
#         result = GetTestResultOneSession(user_id_pw, single_session_id, TestCaseID)
#         if result != 'NT':
#             tested_session_id = single_session_id
#             break
#     if result != 'NT':
#         return result + '(session : ' + str(tested_session_id) +')'
#     return result