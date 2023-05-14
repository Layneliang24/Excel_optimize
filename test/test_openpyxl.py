import os
import re
import time
import openpyxl

path = "C:\\华盛通\\技术支持周报"
os.chdir(path)
main_workbook = openpyxl.load_workbook('线上问题处理记录汇总表20230223.xlsx')
main_workbook.create_sheet('fuck')
main_workbook.save('线上问题处理记录汇总表20230223.xlsx')
main_workbook.close()

