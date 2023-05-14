import datetime
import xlwings as xw
import os
import re
import calendar
from xlwings import constants
from sum.data_transfer import get_addup_start_cell


def find_sheets(wb):
    sheet_list = wb.sheets
    pattern_consult = re.compile(r'.*咨询.*', flags=re.I)
    pattern_unclosed_bug = re.compile(r'.*未关闭.*', flags=re.I)
    pattern_all_bug = re.compile(r'.*某月份.*', flags=re.I)
    pattern_year_sum_up = re.compile(r'.*问题汇总.*', flags=re.I)
    pattern_online = re.compile(r'.*线上问题.*', flags=re.I)
    pattern_handling = re.compile(r'.*正在处理.*', flags=re.I)
    consult, all_bug, unclosed_bug, year_sum_up, online, handling = '', '', '', '', '', ''
    for i in sheet_list:
        if re.search(pattern_consult, string=i.name):
            consult = i
        elif re.search(pattern_all_bug, string=i.name):
            all_bug = i
        elif re.search(pattern_unclosed_bug, string=i.name):
            unclosed_bug = i
        elif re.search(pattern_year_sum_up, string=i.name):
            year_sum_up = i
        elif re.search(pattern_online, string=i.name):
            online = i
        elif re.search(pattern_handling, string=i.name):
            handling = i
    return consult, all_bug, unclosed_bug, year_sum_up, online, handling


# 获取有效的行数，避免中间出现空行
def get_valid_row_count(sheet):
    k = 0
    count = sheet.used_range.last_cell.row
    print('《{}》——期初最后一行的行号{}'.format(sheet, count))
    for i in range(1, sheet.used_range.last_cell.row):
        # 判断空行
        if not sheet.cells(i, 2).value and not sheet.cells(i, 3).value and not sheet.cells(i, 4).value:
            sheet.range('A' + str(i)).api.EntireRow.Clear()  # 清除内容
            print('{}——清除{}行成功'.format(sheet, i))
            k += 1
    if k:
        print('《{}》——需清除空行：{}行'.format(sheet.name, k))
    else:
        print('《{}》——不需要清除空行'.format(sheet.name))
    count = sheet.used_range.last_cell.row
    print('《{}》——期末最后一行的行号{}'.format(sheet, count))
    valid_row_count = count
    return valid_row_count


class Modification(object):
    def __init__(self, target_path, unclosed_path, all_path, wo_path):
        self.app = xw.App(visible=False, add_book=False)  # 程序可见，只打开不新建工作薄
        self.app.display_alerts = False  # 警告关闭
        self.app.screen_updating = False  # 屏幕更新关闭
        self.target_path = target_path
        try:
            self.wb = self.app.books.open(self.target_path)
            self.sheet_consult, self.sheet_all_image, self.sheet_unclosed_image, self.year_sum_up, self.online, self.handling = find_sheets(self.wb)
            self.unclosed_path = unclosed_path
            self.all_path = all_path
            self.wo_path = wo_path
            self.unmerge_cell(self.wb.sheets[3])  # 拆分当月线上问题表格。
            self.fill_num(self.sheet_unclosed_image)
            self.fill_num(self.sheet_all_image)
            self.expand_rows(self.wb.sheets[1], self.sheet_unclosed_image)
            self.expand_rows(self.wb.sheets[3], self.sheet_all_image)
            self.clear_date()
            # self.sum_up_in_year()
        except Exception as e:
            print(e)
            self.app.quit()
            raise Exception(str(e))
        # self.save()

    # 拆分合并单元格，针对”当月线上问题表格“
    @staticmethod
    def unmerge_cell(sheet):
        for i in range(1, sheet.used_range.last_cell.row):
            if sheet.range('N' + str(i)).value:
                sheet.range('N' + str(i)).api.UnMerge()
                print('N' + str(i) + '单元格已经拆分!')

    # 填充序号
    @staticmethod
    def fill_num(source_sheet):
        print('=' * 40 + '填充序号：《' + str(source_sheet.name) + '》' + '=' * 40)
        valid_row_count = get_valid_row_count(source_sheet)
        for i in range(1, valid_row_count+1):
            if source_sheet.cells(i, 1).value != i and source_sheet.cells(i, 2).value:    # 序号不等，而且本行非空
                source_sheet.cells(i, 1).value = i
                print('填充第' + str(i) + '行序号完成')

    # 下拉“正在处理”和“3月线上问题表格”
    # 宏样式
    # Range("A29:I29").Select
    # Selection.AutoFill Destination:=Range("A29:I31"), Type:=xlFillDefault
    @staticmethod
    def expand_rows(target_sheet, source_sheet):
        print('=' * 20 + '下拉列表：《' + str(target_sheet.name) + '》' + '=' * 20)
        target_rows_count = get_valid_row_count(target_sheet) - 1  # 目标表有表头
        source_rows_count = get_valid_row_count(source_sheet)
        source_columns_count = source_sheet.used_range.last_cell.column
        # target_columns_count = target_sheet.used_range.last_cell.column
        print('目标表总行数：' + str(target_rows_count))
        print('源表总行数：' + str(source_rows_count))
        delta = source_rows_count - target_rows_count
        if delta > 0:
            # 扩展首列序号列
            source_range = target_sheet.range('A2:A3').api  # 参考范围
            # 下拉范围，以源表的最大行数 + 1，因为原表没有表头
            fill_range = target_sheet.range('A2:A' + str(source_rows_count + 1)).api
            auto_fill_type = 0
            source_range.AutoFill(fill_range, auto_fill_type)
            # 扩展除了序号之外的所有列
            for i in range(2, source_columns_count + 1):
                target_sheet.range(chr(i + 64) + '2' + ':' + chr(i + 64) + str(source_rows_count + 1)).api.FillDown()
                print('-扩展' + chr(i + 64) + '列完成')
            print('——扩展' + str(delta) + '行完成！')
        else:
            print('不需要下拉表格！')

    # 清除《N月线上问题》无用的日期
    def clear_date(self):
        sheet = self.online
        print('=' * 20 + '清除无效日期：《' + str(sheet.name) + '》' + '=' * 20)
        for cell in sheet.range('I2').expand('down'):
            try:
                if cell.value:
                    if cell.value.strftime("%Y-%m-%d") < '2000-01-00':   # 有值，且值离谱
                        cell.api.ClearContents()
                        cell.offset(column_offset=1).api.ClearContents()
                        print('{}行，清除日期完成!'.format(cell.row))
                else:
                    print('空值')
            except TypeError as e:
                print(e)
                print('{}日期为空'.format(cell.address))

    # ===========================修改《图表统计表格》================================

    # ===========================修改《数据统计表格》================================

    # ===========================修改《年度问题汇总表格》================================
    def sum_up_in_year(self):
        wb = self.wb
        source_sheet = self.online
        target_sheet = self.year_sum_up
        today = datetime.datetime.today()
        this_thursday = today + datetime.timedelta(days=3 - today.weekday())
        last_thursday = this_thursday - datetime.timedelta(days=7)
        this_friday = today + datetime.timedelta(days=4 - today.weekday())
        last_friday = this_friday - datetime.timedelta(days=7)
        print('--- this thursday : {}'.format(this_thursday.strftime("%Y-%m-%d")))
        print('--- last friday : {}'.format(last_friday.strftime("%Y-%m-%d")))
        start_address = source_sheet.range('F2')  # 设置起始单元格为F2，不针对哪一个表
        for i in source_sheet.range('H2').expand('down'):
            # 大于上周四的时间，就是上周五的时间。本周问题就是上周五到本周四。
            if i.value.strftime("%Y-%m-%d") > last_thursday.strftime("%Y-%m-%d"):
                start_address = source_sheet.range('F' + str(i.row))
                break
        # string = chr(start_address.column + 64) + str(start_address.row)
        print('需要复制开始行：{}'.format(start_address.row))
        target_sheet_start_cell = get_addup_start_cell(self.year_sum_up)
        # 跨表格复制内容
        source_sheet.range('A{}:K{}'.format(start_address.row, get_addup_start_cell(source_sheet).row)).api.Copy()
        target_sheet.range("A{}".format(target_sheet_start_cell.row)).api.Select()
        # 粘贴值和数据格式
        target_sheet.api.PasteSpecial(xw.constants.PasteType.xlPasteAllUsingSourceTheme)
        target_sheet.api.PasteSpecial(xw.constants.PasteType.xlPasteValuesAndNumberFormats)
        # target_sheet.api.PasteSpecial(xw.constants.PasteType.xlPasteValues)   # 无效方法
        wb.app.api.CutCopyMode = False
        print('复制完成！')

    # @staticmethod
    # def get_addup_start_cell(sheet):
    #     start_cell = sheet.cells(sheet.used_range.last_cell.row + 1, 1)
    #     print('未使用起始行：', start_cell.row)
    #     add_up_start_cell = start_cell
    #     # 检验是否存在空行
    #     for i in range(1, start_cell.row - 1):
    #         if not sheet.cells(i, 1).value and not sheet.cells(i, 2).value and not sheet.cells(i, 3).value:
    #             add_up_start_cell = sheet.cells(i, 2)
    #             break
    #     return add_up_start_cell

    # ===========================================================
    def save(self):
        self.wb.save(self.target_path)
        self.wb.close()
        self.app.quit()
        print(self.target_path + '——保存修改成功1')


# 修改”图表统计“
class TBTJ(object):
    def __init__(self, target_path, unclosed_path, all_path, wo_path):
        self.app = xw.App(visible=False, add_book=False)  # 程序可见，只打开不新建工作薄
        self.app.display_alerts = False  # 警告关闭
        self.app.screen_updating = False  # 屏幕更新关闭
        self.target_path = target_path
        try:
            self.wb = self.app.books.open(self.target_path)
            self.unclosed_path = unclosed_path
            self.all_path = all_path
            self.wo_path = wo_path
            self.sheet_consult, self.sheet_all_image, self.sheet_unclosed_image, self.year_sum_up, self.online, self.handling = find_sheets(self.wb)
            self.sheet = self.wb.sheets['图表统计']
            self.clean_summary()
            self.modify_chart_title()
        except Exception as e:
            print(e)
            self.app.quit()
            raise Exception(str(e))

    # 修改上方的的小表格
    def clean_summary(self):
        wb = self.app.books.open(self.wo_path)
        sheet = wb.sheets['中间表']
        self.sheet.range('F3').value = str(sheet.used_range.rows.count)
        print('本周线上技术问题数量：{}'.format(str(sheet.used_range.rows.count)))
        self.sheet.range('F4').value = str(get_valid_row_count(self.sheet_consult))
        print('本月线上技术问题数量：{}'.format(str(get_valid_row_count(self.sheet_consult) - 1)))
        wb.close()

    # 修改图表的标题
    def modify_chart_title(self):
        today = datetime.datetime.today()
        this_thursday = today + datetime.timedelta(days=3 - today.weekday())
        this_friday = today + datetime.timedelta(days=4 - today.weekday())
        last_friday = this_friday - datetime.timedelta(days=7)
        print('--- last friday : {}'.format(last_friday.strftime("%Y-%m-%d")))
        week_title = '{}~{}出现的问题统计'.format(last_friday.strftime('%m-%d'), this_thursday.strftime('%m-%d'))
        month_title = '本月出现的问题统计（{}月）'.format(today.strftime('%m'))
        self.sheet.charts['图表 12'].api[1].ChartTitle.Text = week_title
        print('周图表标题：”{}“'.format(week_title))
        self.sheet.charts['图表 13'].api[1].ChartTitle.Text = month_title
        print('月图表标题：“{}”'.format(month_title))

    # 保存
    def save(self):
        self.wb.save(self.target_path)
        self.wb.close()
        self.app.quit()
        print(self.target_path + '——保存修改成功2')


# 修改”数据统计-旧“
class SJTJ_old(object):
    def __init__(self, target_path, unclosed_path, all_path, wo_path):
        self.app = xw.App(visible=False, add_book=False)  # 程序可见，只打开不新建工作薄
        self.app.display_alerts = False  # 警告关闭
        self.app.screen_updating = False  # 屏幕更新关闭
        self.target_path = target_path
        try:
            self.wb = self.app.books.open(self.target_path)
            self.unclosed_path = unclosed_path
            self.all_path = all_path
            self.wo_path = wo_path
            self.sheet_consult, self.sheet_all_image, self.sheet_unclosed_image, self.year_sum_up, self.online, self.handling = find_sheets(self.wb)
            self.sheet = self.wb.sheets['数据统计-旧']
            self.modify_week_summary()
            self.modify_month_summary()
        except Exception as e:
            print(e)
            self.app.quit()
            raise Exception(str(e))

    # 获取第一天和最后一天
    @staticmethod
    def get_first_and_last_day(year, month):
        # 获取当前月的第一天的星期和当月总天数
        weekDay, monthCountDay = calendar.monthrange(year, month)
        # 获取当前月份第一天
        firstDay = datetime.date(year, month, day=1)
        # 获取当前月份最后一天
        lastDay = datetime.date(year, month, day=monthCountDay)
        # 返回第一天和最后一天
        return firstDay, lastDay

    # 修改本周问题统计
    def modify_week_summary(self):
        sheet = self.online  # 线上问题表
        today = datetime.datetime.today()
        this_thursday = today + datetime.timedelta(days=3 - today.weekday())
        last_thursday = this_thursday - datetime.timedelta(days=7)
        this_friday = today + datetime.timedelta(days=4 - today.weekday())
        last_friday = this_friday - datetime.timedelta(days=7)
        print('--- this thursday : {}'.format(this_thursday.strftime("%Y-%m-%d")))
        print('--- last friday : {}'.format(last_friday.strftime("%Y-%m-%d")))
        start_address = sheet.range('F2')  # 设置起始单元格为F2，不针对哪一个表
        for i in sheet.range('H2').expand('down'):
            # 大于上周四的时间，就是上周五的时间。本周问题就是上周五到本周四。
            if i.value.strftime("%Y-%m-%d") > last_thursday.strftime("%Y-%m-%d"):
                start_address = sheet.range('F' + str(i.row))
                break
        string = chr(start_address.column + 64) + str(start_address.row)
        print('本周问题开始单元格：{}'.format(string))
        pattern = re.compile(r'".*"', flags=re.I)
        for i in range(2, 25):  # I2到I24的公式需要设置
            content = re.search(pattern, string=self.sheet.range('I' + str(i)).formula)
            print('I' + str(i) + '公式改为：”{}“'.format('=COUNTIF(\'{}\'!{}:F100,{})' \
                                                    .format(self.online.name, string, content.group())))
            self.sheet.range('I' + str(i)).formula = '=COUNTIF(\'{}\'!{}:F100,{})' \
                .format(self.online.name, string, content.group())
        print('本周问题统计设置完成')

    # 修改本月问题统计表格，严格遵守月日期。
    def modify_month_summary(self):
        sheet = self.online  # 线上问题表
        # 获取当前年份
        year = datetime.date.today().year
        # 获取当前月份
        month = datetime.date.today().month
        first_day = self.get_first_and_last_day(year, month)[0]
        last_day = self.get_first_and_last_day(year, month)[1]
        print('--- first day of this month : {}'.format(first_day.strftime("%Y-%m-%d")))
        print('--- last day of this month : {}'.format(last_day.strftime("%Y-%m-%d")))
        start_address = sheet.range('F2')  # 设置起始单元格为N2，不针对哪一个表
        for i in sheet.range('H2').expand('down'):
            if i.value.strftime("%Y-%m-%d") >= first_day.strftime("%Y-%m-%d"):
                start_address = sheet.range('F' + str(i.row))
                break
        string = chr(start_address.column + 64) + str(start_address.row)
        print('本月问题开始单元格：{}'.format(string))
        pattern = re.compile(r'".*"', flags=re.I)
        for i in range(2, 25):  # N2到N24的公式需要设置
            content = re.search(pattern, string=self.sheet.range('N' + str(i)).formula)
            print('N' + str(i) + '公式改为：”{}“'.format('=COUNTIF(\'{}\'!{}:F100,{})' \
                                                    .format(self.online.name, string, content.group())))
            self.sheet.range('N' + str(i)).formula = '=COUNTIF(\'{}\'!{}:F100,{})' \
                .format(self.online.name, string, content.group())
        print('本月问题统计设置完成')

    def save(self):
        self.wb.save(self.target_path)
        self.wb.close()
        self.app.quit()
        print(self.target_path + '——保存修改成功3')


if __name__ == '__main__':
    path = "C:\\华盛通\\技术支持周报"
    os.chdir(path)

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

    mod = Modification('线上问题处理记录汇总表20230505.xlsx', '线上问题记录-未关闭Bug（3）.xlsx', '线上问题记录-所有Bug（2）.xlsx', '工单查询 (5).xlsx')
    mod.save()
    tbtj = TBTJ('线上问题处理记录汇总表20230505.xlsx', '线上问题记录-未关闭Bug（3）.xlsx', '线上问题记录-所有Bug.xlsx（2）', '工单查询 (5).xlsx')
    tbtj.save()
    sjtj = SJTJ_old('线上问题处理记录汇总表20230505.xlsx', '线上问题记录-未关闭Bug（3）.xlsx', '线上问题记录-所有Bug（2）.xlsx', '工单查询 (5).xlsx')
    sjtj.save()
