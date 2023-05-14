import xlwings as xw
import os
import re
from colorama import Fore
from Excel_optimize.settings import *

# os.chdir(basic_path)


def get_addup_start_cell(sheet):
    start_cell = sheet.cells(sheet.used_range.last_cell.row + 1, 1)
    print('未使用起始行：', start_cell.row)
    add_up_start_cell = start_cell
    # 检验是否存在空行
    for i in range(1, start_cell.row - 1):
        if not sheet.cells(i, 1).value and not sheet.cells(i, 2).value and not sheet.cells(i, 3).value:
            add_up_start_cell = sheet.cells(i, 1)
            break
    return add_up_start_cell


def set_format(sheet_name):
    # sheet_name.range('A1').expand('right').api.autofit()
    sheet_name.range('A1', sheet_name.used_range.last_cell).row_height = 15
    sheet_name.range('A1', sheet_name.used_range.last_cell).api.Font.Size = 12  # 设置字体的大小。
    sheet_name.range('A1', sheet_name.used_range.last_cell).api.Font.Name = 'Calibris'

    # 水平对齐方式：-4108位水平居中，-4131为靠左对齐，-4152为靠右对齐
    sheet_name.range('A1', sheet_name.used_range.last_cell).api.HorizontalAlignment = -4131
    # 垂直对齐方式：-4108位垂直居中，-4160位靠上对齐，-4107为靠下对齐
    sheet_name.range('A1', sheet_name.used_range.last_cell).api.VerticalAlignment = -4130

    # 设置边框线：
    # Borders参数：1为左边框，2为右边框，3为上边框，4为下边框，5为左上至右下对角线，6为右上至左下对角线
    # LinStyle值：1为直线，2为虚线，4为点划线，5为双点划线
    # Weight值：边框线粗细
    sheet_name.range('A1', sheet_name.used_range.last_cell).api.Borders(1).LineStyle = 7
    sheet_name.range('A1', sheet_name.used_range.last_cell).api.Borders(1).Weight = 1

    sheet_name.range('A1', sheet_name.used_range.last_cell).api.WrapText = False
    sheet_name.range('A1', sheet_name.used_range.last_cell).api.Orientation = 0
    sheet_name.range('A1', sheet_name.used_range.last_cell).api.AddIndent = False
    sheet_name.range('A1', sheet_name.used_range.last_cell).api.IndentLevel = 0
    sheet_name.range('A1', sheet_name.used_range.last_cell).api.ShrinkToFit = False
    sheet_name.range('A1', sheet_name.used_range.last_cell).api.MergeCells = False
    # sheet_name.range('A1').expand('table').api.Style = 'Normal'

    info = '''
    行高：{}
    字体大小：{}
    设置字体：{}
    水平居中：{}
    垂直居中：{}
    '''.format(sheet_name.range('A1').expand('table').row_height,
               sheet_name.range('A1').expand('table').api.Font.Size,
               sheet_name.range('A1').api.Font.Name,
               sheet_name.range('A1').expand('table').api.HorizontalAlignment,  # -4108 水平居中。 -4131 靠左，-4152 靠右。
               sheet_name.range('A1').expand('table').api.VerticalAlignment
               )
    print(info)
    print('——设置格式完成')


class Transfer(object):
    def __init__(self, target_path, unclosed_path, all_path, wo_path):
        self.app = xw.App(visible=False, add_book=False)  # 程序可见，只打开不新建工作薄
        self.app.display_alerts = False  # 警告关闭
        self.app.screen_updating = False  # 屏幕更新关闭
        try:
            self.target_path = target_path
            self.wb = self.app.books.open(self.target_path)
            self.sheet_consult, self.sheet_all, self.sheet_unclosed = self.find_sheets()
            self.unclosed_path = unclosed_path
            self.all_path = all_path
            self.wo_path = wo_path
            self.flag_list = [i.value[0] for i in self.sheet_consult.range('A1').expand('table').rows]  # 标志符列表
            self.bug_2_target(self.unclosed_path, self.sheet_unclosed)
            self.bug_2_target(self.all_path, self.sheet_all)
            self.wo_2_target(self.wo_path, self.sheet_consult)
            # 设置格式
            set_format(self.sheet_consult)
            set_format(self.sheet_unclosed)
            set_format(self.sheet_all)
            self.save()
        except Exception as e:
            print(e)
            self.app.quit()
            raise Exception(str(e))

    # 判断本周时间是否跨月了，
    # 1.如果本周都是月中，则不需要清空咨询汇总，只需要拼接起来，不需要修改表名
    # 2.如果本周跨月了，不需要清空咨询汇总，需要将表名改为N~N+1月。但需要在下周清空咨询表。
    # 3.如果是上周跨月、或者本周周期刚好是新月期了，则本周需要清空咨询汇总表，并且更改表名
    # 比如如果上周都是3月，本周都是4月，也需要清空，正好没跨月怎么处理？8.25~8.31是上周，9.1~9.7是本周
    # 判断上周五和本周五是否同月，如果不同就需要清空，并且改表名。
    @staticmethod
    def judge_clean():
        today = datetime.datetime.today()
        # 周一为第一天
        # weekday，返回一个整数代表星期几，0表示星期一，6表示星期日。
        this_monday = today - datetime.timedelta(days=today.weekday())
        # this_sunday = today + datetime.timedelta(days=6 - today.weekday())
        this_friday = today + datetime.timedelta(days=4 - today.weekday())
        this_thursday = today + datetime.timedelta(days=3 - today.weekday())
        last_friday = this_friday - datetime.timedelta(days=7)
        last_thursday = this_thursday - datetime.timedelta(days=7)
        # last_monday = this_monday - datetime.timedelta(days=7)
        # last_last_thursday = this_thursday - datetime.timedelta(days=14)
        last_last_friday = this_friday - datetime.timedelta(days=14)
        print('周报表日期：{}至{}'.format(last_friday.strftime('%Y-%m-%d'), this_thursday.strftime('%Y-%m-%d')))
        # 如果本轮周期是月中，同时上轮周期是月中，则啥也不用改，
        if this_thursday.strftime("%Y-%m") == last_friday.strftime("%Y-%m") and \
                last_last_friday.strftime("%Y-%m") == last_thursday.strftime("%Y-%m") == this_monday.strftime("%Y-%m"):
            print('本轮周期是月中，不需要清空表、不更改表名')
            return False
        # 如果本轮跨月，则需要改表名
        elif this_thursday.strftime("%Y-%m") > last_friday.strftime("%Y-%m"):
            print('本轮周期跨月，不需要清空表，需要更改表名')
            return False
        elif last_thursday.strftime("%Y-%m") > last_last_friday.strftime("%Y-%m") or \
                last_friday.strftime("%Y-%m") == this_thursday.strftime("%Y-%m"):
            print('上轮周期跨月或者本轮周期全是新月，需要清空表，并且修改表名')
            return True

    def find_sheets(self):
        sheet_list = self.wb.sheets
        pattern_consult = re.compile(r'.*咨询.*', flags=re.I)
        pattern_unclosed_bug = re.compile(r'.*未关闭.*', flags=re.I)
        pattern_all_bug = re.compile(r'.*某月份.*', flags=re.I)
        consult, all_bug, unclosed_bug = '', '', ''
        for i in sheet_list:
            if re.search(pattern_consult, string=i.name):
                consult = i
            elif re.search(pattern_all_bug, string=i.name):
                all_bug = i
            elif re.search(pattern_unclosed_bug, string=i.name):
                unclosed_bug = i
        return consult, all_bug, unclosed_bug

    # 复制”未关闭bug“，“所有bug”到目标表格
    def bug_2_target(self, source_path, target_sheet):
        # 清空目标表格
        target_sheet.range('A1').expand('table').clear()
        print('清空表格完成!')
        print(Fore.RED + '=' * 10 + '复制《' + str(source_path) + '》到——>  ' + str(target_sheet) + '=' * 10)
        sheet_value = []
        wb = self.app.books.open(source_path)
        sheet = wb.sheets['Bug']
        for j in range(2, sheet.used_range.rows.count + 1):
            single_row_value = sheet.range(
                'A' + str(j) + ':' + chr(sheet.used_range.columns.count + 64) + str(j)).value  # ('A2:A5')
            print(Fore.CYAN + '第' + str(j - 1) + '行：' + str(single_row_value))
            sheet_value.append(single_row_value)
        # A1那一列需要用来填充序号
        target_sheet.range('B1').expand('table').value = sheet_value
        print('复制——' + source_path + '完成' + '，总共' + str(sheet.used_range.rows.count - 1) + '行')
        wb.close()
        self.wb.save(self.target_path)

    '''
    def copy_2_target_2(self, source_path, target_sheet):
        print('=' * 40 + '复制《' + str(source_path) + '》' + '=' * 40)
        wb = self.app.books.open(source_path)
        sheet = wb.sheets['Bug']
        sheet.range('A2').copy()
        target_sheet.range('B1').paste()
        print('复制——' + source_path + '完成' + '，总共' + str(sheet.used_range.rows.count - 1) + '行')
        wb.close()
        self.wb.save(self.target_path)
    '''

    # 复制工单表格到目标表格
    def wo_2_target(self, source_path, target_sheet):
        if self.judge_clean():
            target_sheet.range('A1').expand('table').clear()
            self.flag_list = []  # 如果需要清空表格，那么标志位也必须清空
        print('=' * 20 + '复制《' + str(source_path) + '》' + '=' * 20)
        sheet_value = []
        wb = self.app.books.open(source_path)
        sheet = wb.sheets['中间表']
        for j in range(2, sheet.used_range.rows.count + 1):
            # ('A2:A5')
            single_row_value = sheet.range('A' + str(j) + ':' + chr(sheet.used_range.columns.count + 64) + str(j)).value
            print('第' + str(j - 1) + '行：' + str(single_row_value))
            if single_row_value[0] in self.flag_list:
                print('——数据已存在，跳过')
                continue
            sheet_value.append(single_row_value)
        target_sheet.range(get_addup_start_cell(self.sheet_consult)).expand('table').value = sheet_value
        print('复制——' + source_path + '完成' + '，总共' + str(sheet.used_range.rows.count - 1) + '行')
        wb.close()
        self.wb.save(self.target_path)

    def save(self):
        self.wb.save(self.target_path)
        self.wb.close()
        self.app.quit()
        print(self.target_path + '——保存修改成功')


if __name__ == '__main__':
    tran = Transfer('线上问题处理记录汇总表20230330.xlsx', '线上问题记录-未关闭Bug.xlsx', '线上问题记录-所有Bug.xlsx', '工单查询 (1).xlsx')
