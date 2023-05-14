import time
import os
import xlwings as xw

app = xw.App(visible=True, add_book=False)  # 程序可见，只打开不新建工作薄
app.display_alerts = True  # 警告关闭
app.screen_updating = True  # 屏幕更新打开
path = 'C:/华盛通/技术支持周报/线上问题处理记录汇总表20230302.xlsx'
workbook = app.books.open(os.path.abspath(path))
source_sheet = workbook.sheets['未关闭映射模板']
cell = source_sheet.used_range.last_cell
print(cell.row, cell.column)
rows = cell.row
columns = cell.column
list = source_sheet.range('A2', (rows, columns)).value
print(list)
workbook.save()
workbook.close()
app.quit()
