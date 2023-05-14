import xlwings as xw
import re
import os
from Excel_optimize.settings import *


# 输出信息
def print_info(book):
    print('=' * 20 + '《' + book + '》' + '=' * 20)


# 获取表格的最大行数数，作为循环的依据，更为可靠。
def get_max_row(sheet):
    valid_row_count = sheet.used_range.last_cell.row
    for i in range(1, sheet.used_range.last_cell.row + 1):
        # 判断依据是连续三个单元格都是空的话，那么这行就是空行，上一行就是最大的非空有效行。
        if not sheet.cells(i, 1).value and not sheet.cells(i, 2).value and not sheet.cells(i, 3).value:
            valid_row_count = i - 1
    print('《{}》——最大有效行行号{}'.format(sheet.name, valid_row_count))
    return valid_row_count


# 清理未关闭bug以及所有bug这两个表格
class CleanBugExcel(object):
    def __init__(self, book_path, sheet_name):
        self.app = xw.App(visible=True, add_book=False)  # 程序可见，只打开不新建工作薄
        self.app.display_alerts = False  # 警告关闭
        self.app.screen_updating = True  # 屏幕更新关闭
        self.book_path = book_path
        print('整理日期：' + str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        print_info(self.book_path)
        try:
            self.wb = self.app.books.open(self.book_path)
            self.sheet_name = sheet_name
            self.sheet = self.wb.sheets[self.sheet_name]
            self.name_list = target_column
            self.start_cell_priority = self.get_column_name()[self.name_list[0]]
            self.start_cell_status = self.get_column_name()[self.name_list[1]]
            self.start_cell_assigner = self.get_column_name()[self.name_list[2]]
            self.start_cell_date = self.get_column_name()[self.name_list[3]]
            self.modify_bug_status()
            self.modify_assigner()
            self.modify_priority()
            self.key = self.get_column_name()['创建日期']  # 按创建日期排序
            self.sort_date()
            self.save()
        except Exception as e:
            print(e)
            self.app.quit()
            raise Exception(str(e))

    # 获取表格的最大行数数，作为循环的依据，更为可靠。
    @staticmethod
    def get_max_row(sheet):
        valid_row_count = sheet.used_range.last_cell.row
        for i in range(1, sheet.used_range.last_cell.row + 1):
            # 判断依据是连续三个单元格都是空的话，那么这行就是空行，上一行就是最大的非空有效行。
            if not sheet.cells(i, 1).value and not sheet.cells(i, 2).value and not sheet.cells(i, 3).value:
                valid_row_count = i - 1
        print('《{}》——最大有效行行号{}'.format(sheet.name, valid_row_count))
        return valid_row_count

    # 过滤字符串，比如过滤掉”新加坡“，”马来西亚“，”新西兰“等的Bug
    def filter_string(self):
        pass

    # 获取表头
    def get_column_name(self):
        head_list = self.sheet.range('A1').expand('right')
        start_cell_dict = {}
        for i in self.name_list:
            for cell in head_list:
                if cell.value == i:
                    # 通过字典方式存储表头单元对象
                    start_cell_dict[i] = self.sheet.cells(cell.row + 1, cell.column)
        return start_cell_dict

    # 修改”Bug状态“
    def modify_bug_status(self):
        end_cell_status = self.sheet.cells(self.get_max_row(self.sheet), self.start_cell_status.column)
        target = self.sheet.range(self.start_cell_status, end_cell_status)  # 必须是某个确定的单元格，不能写成F
        for i in target:
            try:
                i.value
            except TypeError:
                print('单元格为空，请填充！')
                continue
            else:
                if i.value == '激活':
                    i.value = '处理中'
                    print('修改Bug状态：' + '第' + str(i.row) + '行' + '从“激活”改为——>“处理中”')
                elif i.value == '已解决':
                    i.value = '验证中'
                    print('修改Bug状态：' + '第' + str(i.row) + '行' + '从“已解决”改为——>“验证中”')
        print(self.book_path + '——修改Bug状态完成')

    # 修改”指派给“
    def modify_assigner(self):
        pattern = re.compile(r'\(.*?\)', flags=re.I)  # 正则表达式去除括号及括号内的内容
        end_cell_assigner = self.sheet.cells(self.get_max_row(self.sheet), self.start_cell_assigner.column)
        target = self.sheet.range(self.start_cell_assigner, end_cell_assigner)
        for i in target:
            try:
                i.value = re.sub(pattern, '', i.value)
            except TypeError:
                print('第{}行单元格为空，请填充！！！'.format(i))
            else:
                print('简化指派人：' + '第' + str(i.row) + '行' + '改为——>' + '"' + i.value + '"')
                if self.sheet.range("J" + str(i.row)).value:
                    i.value = self.sheet.range("J" + str(i.row)).value
                    print('修改指派人：' + '第' + str(i.row) + '行' + '从“Closed”改为——>' + '"' + self.sheet.range(i.row,
                                                                                                        10).value + '"')
        print(self.book_path + '——清洗指派人完成')

    # 修改“优先级”
    def modify_priority(self):
        end_cell_priority = self.sheet.cells(self.get_max_row(self.sheet), self.start_cell_priority.column)
        target = self.sheet.range(self.start_cell_priority, end_cell_priority)
        for i in target:
            if i.value == '中':
                i.value = '普通'
                print('修改优先级：' + '第' + str(i.row) + '行，从”中“改为”普通“')
        print(self.book_path + '——清洗优先级完成')

    # 把日期改为升序
    def sort_date(self):
        end_cell_date = self.sheet.cells(self.get_max_row(self.sheet), self.start_cell_date.column)
        target = self.sheet.range(self.start_cell_date, end_cell_date)
        # last_cell确保所有列都排序
        self.sheet.range('A2', self.sheet.used_range.last_cell).api.Sort(Key1=self.sheet.range(self.key).api, Order1=1,
                                                                         Orientation=1)
        for i in target:
            print('排序：第' + str(i.row) + '行的日期：' + i.value)
        print(self.book_path + '——排序完成')

    # 保存修改
    def save(self):
        self.wb.save(self.book_path)
        self.wb.close()
        self.app.quit()
        print(self.book_path + '——保存修改成功')


# 清理工单表格
class CleanWOExcel(object):
    def __init__(self, *book_path):
        self.app = xw.App(visible=True, add_book=False)  # 程序可见，只打开不新建工作薄
        self.app.display_alerts = False  # 警告关闭
        self.app.screen_updating = True  # 屏幕更新关闭
        print('整理日期：' + str(modify_time))
        try:
            self.book_path_target = book_path[0]  # 用第一个工单表格
            print_info(self.book_path_target)
            self.wb = self.app.books.open(self.book_path_target)
            self.sheet = self.wb.sheets['第一页']
            self.flag_list = [i.value[0] for i in self.sheet.range('A3').expand('table').rows]  # 标志符列表
            self.wo_list = [i for i in book_path]
            self.last_cell = self.sheet.used_range.last_cell
            self.add_up_start_cell = self.sheet.range(get_max_row(self.sheet) + 1, 1)
            self.init_rows = get_max_row(self.sheet)
            self.add_up()
            self.create_middle_sheet()
            self.sort_date()
            self.save()
            # except FileNotFoundError:
            #    self.app.quit()
            #    raise FileNotFoundError('找不到指定文件!')
        except Exception as e:
            print(e)
            self.app.quit()
            raise Exception(str(e))

    # 汇总所有表格内容
    def add_up(self):
        excel_value = []
        start_cell = self.add_up_start_cell
        for i in self.wo_list[1:]:
            print('------正在汇总表格——%s------' % i)
            sheet_value = []
            wb = self.app.books.open(i)
            sheet = wb.sheets['第一页']
            for j in range(3, sheet.used_range.rows.count + 1):
                single_row_value = sheet.range('A' + str(j)).expand('right').value
                print(single_row_value, len(single_row_value), sep='——')
                sheet_value.append(single_row_value)
            excel_value.append(sheet_value)
        self.sheet.range(start_cell).expand('table').value = self.convert_3_to_2_dimension(excel_value)
        print(self.book_path_target + '——汇总完成' + '，新增' + str(len(self.convert_3_to_2_dimension(excel_value))) + '行')
        print('原有行数：' + str(self.init_rows))
        print('现有行数：' + str(get_max_row(self.sheet)))
        print(self.book_path_target + '——汇总完成')

    # 新建一个中间表存储更改后的内容
    def create_middle_sheet(self):
        print('=' * 40 + '新建表格到——《' + self.book_path_target + '》' + '=' * 40)
        if '中间表' not in [i.name for i in self.wb.sheets]:
            self.wb.sheets.add('中间表')
        middle_sheet = self.wb.sheets['中间表']
        middle_sheet.clear()  # 清空表
        middle_sheet.range('A1:H1').value = ['咨询内容', '问题类型', '业务类型', '细分类型', '咨询人', '部门', '咨询时间', '处理人']
        max_row = self.sheet.used_range.last_cell.row
        target_list = [3, 4, 10]  # 要合并的列号
        col_num = [6, 7, 8, 11, 13, 12]  # 要复制的列号
        main_list = [[], [], []]
        new_list = []
        # 把要合并的列的值分别放进二维数组里
        for k in range(len(target_list)):  # 0，1，2
            for i in range(3, max_row + 1):  # 逐行迭代，从源表的第3行开始
                main_list[k].append(str(self.sheet.cells(i, target_list[k]).value).strip())  # 去掉空格
        # 把二维数组合并
        for j in range(0, len(main_list[0])):
            new_list.append(main_list[0][j] + ' ' + main_list[1][j] + ' ' + main_list[2][j])
        # 写入合并的数据到目标表第1列
        for n in range(0, len(new_list)):
            middle_sheet.cells(n + 2, 1).value = new_list[n]
        # 把其余的数据写入目标表
        for r in range(0, len(col_num)):
            for m in range(2, max_row + 1):
                # 目标表的(2,2)的值应该等于源表的(3,n)的值
                middle_sheet.cells(m, r + 2).value = self.sheet.cells(m + 1, col_num[r]).value  # [列，行] = [列的切片值，行]
        # 最后一列的值写“工单”
        for k in range(2, max_row):
            middle_sheet.range('H' + str(k)).value = ['工单']
        self.set_format(middle_sheet)
        self.wb.save(self.book_path_target)

    # 把日期改为升序
    def sort_date(self):
        sheet = self.wb.sheets['中间表']
        target = sheet.range('G2').expand('down')
        cell = self.wb.sheets['中间表'].used_range.last_cell
        sheet.range('A2', cell).api.Sort(Key1=sheet.range('G2').api, Order1=1, Orientation=1)
        for i in target:
            print('排序：第' + str(i.row) + '行的日期：' + str(i.value))
        print(self.book_path_target + '——排序完成')

    # 设置内容格式
    def set_format(self, sheet_name):
        sheet_name.autofit()
        sheet_name.range('A1').expand('down').column_width = 80
        print(self.book_path_target + '——设置格式完成')

    # 把3维列表变成2维列表
    def convert_3_to_2_dimension(self, source_list):
        result_list = []
        for i in source_list:
            for j in i:
                if j[0] in self.flag_list:
                    print('{}/\n——数据已存在，跳过'.format(j))
                    continue
                result_list.append(j)
        return result_list

    # 保存表格
    def save(self):
        self.wb.save(self.book_path_target)
        self.wb.close()
        self.app.quit()
        print(self.book_path_target + '——保存修改成功')


if __name__ == '__main__':
    os.chdir(basic_path)
    clean = CleanBugExcel('线上问题记录-所有Bug.xlsx', 'Bug')
    clean_unclosed = CleanBugExcel('线上问题记录-未关闭Bug.xlsx', 'Bug')
    clean_wo = CleanWOExcel('工单查询 (1).xlsx', '工单查询 (2).xlsx', '工单查询 (3).xlsx')
