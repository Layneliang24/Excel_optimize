from tkinter import *
from tkinter.ttk import *
import ttkbootstrap as ttk
from ttkbootstrap.tooltip import ToolTip
from tkinter import filedialog
import sys, os, re
import tkinter
from tkinter import messagebox
from Excel_optimize import settings
from cle.data_clean import CleanWOExcel, CleanBugExcel
from sum.data_transfer import Transfer
from mod.modify_sheet import Modification, TBTJ, SJTJ_old

"""
全局通用函数
"""


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


# 初始化函数处理每一个目标工作簿，返回字典
# {book_name:工作簿名,sheet_name_list:[表名],X_sheet_max_row:某表的最大有效行}
# 获取到最大有效行数，最后的一个单元格，表头
def init():
    pass


def start(together, unclosed_bug, all_bug, wo_list):
    CleanBugExcel(all_bug, 'Bug')
    CleanBugExcel(unclosed_bug, 'Bug')
    CleanWOExcel(*wo_list)  # 这个方式可以直接将一个列表的所有元素当作不定参数
    Transfer(together, unclosed_bug, all_bug, wo_list[0])  # 顺序不能搞乱
    mod = Modification(together, unclosed_bug, all_bug, wo_list[0])
    mod.save()
    tbtj = TBTJ(together, unclosed_bug, all_bug, wo_list[0])
    tbtj.save()
    sjtj = SJTJ_old(together, unclosed_bug, all_bug, wo_list[0])
    sjtj.save()


class mystdout(object):
    def __init__(self, text):
        self.stdoutbak = sys.stdout
        self.stderrbak = sys.stderr
        sys.stdout = self
        sys.stderr = self
        self.text = text

    def write(self, info):  # 外部的print语句将执行本write()方法，并由当前sys.stdout输出
        # t = tkinter.Text()
        self.text.insert('end', info)
        self.text.update()
        self.text.see(tkinter.END)

    def restore_std(self):
        sys.stdout = self.stdoutbak
        sys.stderr = self.stderrbak


# 自动隐藏滚动条
def scrollbar_auto_hide(bar, widget):
    def show():
        bar.lift(widget)

    def hide():
        bar.lower(widget)

    hide()
    widget.bind("<Enter>", lambda e: show())
    bar.bind("<Enter>", lambda e: show())
    widget.bind("<Leave>", lambda e: hide())
    bar.bind("<Leave>", lambda e: hide())


# 容器类，继承于Frame，本类实现了许多组件的自动挂载
class Frame(ttk.Frame):
    def __init__(self):
        super().__init__()  # 对继承自父类的属性进行初始化，并且用父类的初始化方法初始化继承的属性。
        """
        super()的使用
        python2必须写成 super(子类，self).方法名(参数)
        python3也可直接写成 super().方法名(参数)
        """
        self.locate()  # 调用自身的方法，把自身放置到某个坐标位置。
        label = Label(self, text="报表整合工具", anchor="center")
        label.grid(row=1, column=1, padx=100, pady=20)
        label = Label(self, text="更换皮肤", anchor="center")
        label.grid(row=1, column=3)
        select_skin = self.tk_select_box_sele_skin()
        select_skin.grid(row=1, column=4)
        # 工单
        label = Label(self, text="选择工单报表（多选）", anchor="center")
        label.grid(row=2, column=1, pady=0)
        self.select_path_WO = ttk.StringVar()
        ipt = Entry(self, textvariable=self.select_path_WO, width=30)
        ipt.grid(row=3, column=1, padx=30, pady=0)
        btn = Button(self, text="选择文件", command=self.select_file_wo)
        btn.grid(row=3, column=2)
        # 未关闭报表
        label = Label(self, text="未关闭Bug报表", anchor="center")
        label.grid(row=4, column=1)
        self.select_path_Un = ttk.StringVar()
        ipt = Entry(self, textvariable=self.select_path_Un, width=30)
        ipt.grid(row=5, column=1)
        btn = Button(self, text="选择文件", command=self.select_file_un)
        btn.grid(row=5, column=2)
        # 所有报表
        label = Label(self, text="所有Bug报表", anchor="center")
        label.grid(row=6, column=1)
        self.select_path_All = ttk.StringVar()
        ipt = Entry(self, textvariable=self.select_path_All, width=30)
        ipt.grid(row=7, column=1)
        btn = Button(self, text="选择文件", command=self.select_file_all)
        btn.grid(row=7, column=2)
        # 目标文件
        label = Label(self, text="目标文件", anchor="center")
        label.grid(row=8, column=1)
        self.select_path_Tag = ttk.StringVar()
        ipt = Entry(self, textvariable=self.select_path_Tag, width=30)
        ipt.grid(row=9, column=1)
        btn = Button(self, text="选择文件", command=self.select_file_tag)
        btn.grid(row=9, column=2)
        # 选择文件夹
        label = Label(self, text="选择文件夹", anchor="center")
        label.grid(row=9, column=1)
        self.select_path_folder = ttk.StringVar()
        ipt = Entry(self, textvariable=self.select_path_folder, width=30)
        ipt.grid(row=9, column=1)
        btn = Button(self, text="选择文件夹", command=self.select_folder)
        btn.grid(row=9, column=2)
        # 启动按钮
        btn = Button(self, text="合并", command=self.read_log)  # 注意不能带括号。
        # default tooltip
        ToolTip(btn, text="请不要重复点击！")
        btn.grid(row=10, column=1)
        # 输出框
        self.text = Text(self, borderwidth=5, bg='#B0E0E6', width=80, height=38)
        # self.text.tag_add('tag',)
        # self.text.tag_config(foreground='red')
        self.text.grid(row=2, column=3, rowspan=20, columnspan=10, padx=10, pady=30, sticky=ttk.SE)

    def locate(self):  # 调用自己的方法，把自己放置到某个位置
        self.grid(row=0, column=0, rowspan=20, columnspan=20)

    def tk_select_box_sele_skin(self):
        style = ttk.Style()
        theme_names = style.theme_names()  # 以列表的形式返回多个主题名
        cb = ttk.Combobox(self, values=theme_names)
        cb.current(theme_names.index(theme_names[8]))

        def change_theme(event):
            theme_cbo_value = cb.get()
            style.theme_use(theme_cbo_value)
            cb.selection_clear()

        cb.bind('<<ComboboxSelected>>', change_theme)
        return cb

    # 执行函数
    def select_file_wo(self):
        # 多个文件选择
        selected_files_path = filedialog.askopenfilenames()  # askopenfilenames函数选择多个文件
        self.select_path_WO.set('\n'.join(selected_files_path))  # 多个文件的路径用换行符隔开

    def select_file_un(self):
        # 单个文件选择
        selected_file_path = filedialog.askopenfilename()  # 使用askopenfilename函数选择单个文件
        self.select_path_Un.set(selected_file_path)

    def select_file_all(self):
        # 单个文件选择
        selected_file_path = filedialog.askopenfilename()  # 使用askopenfilename函数选择单个文件
        self.select_path_All.set(selected_file_path)

    def select_file_tag(self):
        # 单个文件选择
        selected_file_path = filedialog.askopenfilename()  # 使用askopenfilename函数选择单个文件
        self.select_path_Tag.set(selected_file_path)

    def select_folder(self):
        # 文件夹选择
        selected_folder = filedialog.askdirectory()  # 使用askdirectory函数选择文件夹
        self.select_path_folder.set(selected_folder)

    def read_log(self):
        # 清空输出框
        if not self.text.compare("end-1c", "==", "1.0"):
            self.text.delete("1.0", END)
        out = mystdout(self.text)  # 实例化了本mystdout类，print输出就换了地方了
        if not self.select_path_folder.get():
            messagebox.showerror(title='出错', message='请先选择操作路径！')
        else:
            try:
                os.chdir(self.select_path_folder.get())
                self.send_folder_path()
                to, un, al, wo = find_books(self.select_path_folder.get())
                start(to, un, al, wo)
            except Exception as e:
                messagebox.showerror(title='出错', message=str(e))

    def send_folder_path(self):
        folder_path = self.select_path_folder.get()
        settings.basic_path = folder_path


# 窗口类，继承于Tk
class WinGUI(ttk.Window):  # 让这个窗口继承于ttk的Window
    def __init__(self):
        super().__init__()
        self.set_attr()
        self.frame = Frame()  # 实例化本窗口类或者他的子类时，就会自动实例化一个容器类Frame，同时这个容器马上就有了各种组件
        self.themename = 'solar'

    def set_attr(self):
        self.title("报表整理工具")
        # 设置窗口大小、居中
        width = 1200
        height = 900
        screenwidth = self.winfo_screenwidth()
        screenheight = self.winfo_screenheight()
        geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 3)
        self.geometry(geometry)
        self.resizable(width=False, height=False)


class Win(WinGUI):
    def __init__(self):
        super().__init__()
        self.event_bind()

    def event_bind(self):
        pass


if __name__ == "__main__":
    win = Win()
    win.mainloop()
