import time
import xlwings as xw

app = xw.App(visible=True, add_book=False)  # 程序可见，只打开不新建工作薄
app.display_alerts = False  # 警告关闭
app.screen_updating = True  # 屏幕更新打开
path = 'C:\\华盛通\\技术支持周报\\线上问题处理记录汇总表20230223.xlsx'
wb = app.books.open(path)


class transfer_to_main_unclosed_bug(object):
    def __init__(self):
        self.sheet = wb.sheets['未关闭映射模板']
        self.sheet.clear_contents()         # 清空内容，但保留格式
        self.source_wb = app.books.open('C:\\华盛通\\技术支持周报\\线上问题记录-未关闭Bug.xlsx')
        self.copy()
        self.final()

    def copy(self):
        self.source_sheet = self.source_wb.sheets['Bug']
        cell = self.source_sheet.used_range.last_cell
        rows = cell.row
        columns = cell.column
        list = self.source_sheet.range('A2', (rows, columns)).value
        print(list)
        self.sheet.range('A1').options(expand='table').value = list
        self.sheet.used_range.columns.autofit()

    def final(self):
        self.source_wb.save()
        self.source_wb.close()


class transfer_to_main_all_bug(object):
    def __init__(self):
        self.sheet = wb.sheets['某月份映射模板']
        self.sheet.clear_contents()
        self.source_wb = app.books.open('C:\\华盛通\\技术支持周报\\线上问题记录-所有Bug.xlsx')
        self.copy()
        self.final()

    def copy(self):
        self.source_sheet = self.source_wb.sheets['Bug']
        list = self.source_sheet['A2:K51'].value
        # for i in list:
        #     print(i)
        self.sheet.range('A1').options(expand='table').value = list
        self.sheet.used_range.columns.autofit()

    def final(self):
        self.source_wb.save()
        self.source_wb.close()


class transer_to_main_consult_content(object):
    def __init__(self):
        self.sheet = wb.sheets['2~3月咨询汇总']
        self.source_wb = app.books.open('C:\\华盛通\\技术支持周报\\总表.xlsx')
        self.shape = self.sheet.used_range.shape
        self.copy()
        self.final()

    def copy(self):
        self.source_sheet = self.source_wb.sheets['新']
        # 如何确定源数据的已使用范围和去除首行:
        cell = self.source_sheet.used_range.last_cell
        rows = cell.row
        columns = cell.column
        list=self.source_sheet.used_range('a2', (rows, columns)).value
        print('确认源数据范围'+ list)
        self.sheet.range('A2').options(expand='table').value = list

    def final(self):
        self.source_wb.save()
        self.source_wb.close()


tr = transfer_to_main_unclosed_bug()
time.sleep(3)
tr1 = transfer_to_main_all_bug()
time.sleep(3)
tr2 = transer_to_main_consult_content()
wb.save()
wb.close()
app.quit()
