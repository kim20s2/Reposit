# coding=utf-8
import os
import re
from enum import Enum
import datetime
import time
import subprocess
import chardet

user_id_pw0 = "kim30s2"
user_id_pw1 = "kim30s2"

view_issue_command = 'im viewissue --user=' + user_id_pw0 + ' --password=' + user_id_pw1 + ' '        # user.txt파일에서 얻은 user id와 pw입력 : im viewissue --user=kim30s2 --password=kim30s2
im_command_txt = view_issue_command + '1228293'                                                       # im viewissue --user=kim30s2 --password=kim30s2 1228293
start = time.time()
result = subprocess.check_output(im_command_txt, shell=True)
end = time.time()
print("get_luk_name함수에서 im viewissue cmd걸린시간" + f"{end - start: .5f} sec")