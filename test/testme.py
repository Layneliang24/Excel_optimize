import re,os

def find_books(path):
    file_list = [i for i in os.walk(path)]  # 返回二维列表[路径，文件夹，文件]
    book_list = file_list[0][2]  # 获取目标目录第一级目录下的表格文件
    pattern_wo = re.compile(r'.*工单.*', flags=re.I)
    pattern_unclosed_bug = re.compile(r'.*未关闭.*', flags=re.I)
    pattern_all_bug = re.compile(r'.*所有.*', flags=re.I)
    pattern_together = re.compile(r'.*汇总.*', flags=re.I)
    wo_list, all_bug, unclosed_bug, together = [], '所有Bug', '未关闭Bug', '汇总表'
    for i in book_list:
        if re.search(pattern_wo, string=i):
            wo_list.append(i)
        elif re.search(pattern_all_bug, string=i):
            all_bug = i
        elif re.search(pattern_unclosed_bug, string=i):
            unclosed_bug = i
        elif re.search(pattern_together, string=i):
            together = i
    return together, unclosed_bug, all_bug, wo_list

path1 = 'C:\华盛通\技术支持周报'
path2 = 'C:\华盛通\技术支持周报\原件-4.6'
path3 = 'C:\华盛通\技术支持周报\历史周报'
print(find_books(path3))