import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog
import tkinter as tk


# 创建窗口
root = ttk.Window()
root.geometry('1000x800+500+100')
root.title('报表整合工具')

# 设置窗口主题
stylee = ttk.Style()
theme_names = stylee.theme_names()  # 以列表的形式返回多个主题名
theme_selection = ttk.Frame(root, padding=(10, 10, 10, 0))
theme_selection.grid(row=0, column=2)
lbl = ttk.Label(theme_selection, text="切换主题:")
lbl.grid(row=0, column=3, sticky=W)
theme_cbo = ttk.Combobox(
    master=theme_selection,
    values=theme_names,
)
theme_cbo.grid(row=0, column=4, sticky=W)
theme_cbo.current(theme_names.index(stylee.theme.name))



def change_theme(event):
    theme_cbo_value = theme_cbo.get()
    stylee.theme_use(theme_cbo_value)
    theme_cbo.selection_clear()


theme_cbo.bind('<<ComboboxSelected>>', change_theme)


class main_interface(object):
    def __init__(self, window):
        self.window = window
        self.frame = ttk.Frame(self.window, borderwidth=1)
        self.frame.grid(row=0, column=2)
        self.select_path = tk.StringVar()
        # 布局控件
        self.lab_workorder = ttk.Label(self.frame, text='请选择工单报表（可以多选）')
        # self.lab_workorder.grid(row=1, column=0)
        self.ent_workorder = tk.Entry(root, textvariable=self.select_path)
        self.ent_workorder.grid(row=2, column=0)
        self.but_workorder = tk.Button(root, text="选择文件", command=self.select_files)
        self.but_workorder.grid(row=2, column=1)

    def select_file(self):
        # 单个文件选择
        selected_file_path = filedialog.askopenfilename()  # 使用askopenfilename函数选择单个文件
        self.select_path.set(selected_file_path)

    def select_files(self):
        # 多个文件选择
        selected_files_path = filedialog.askopenfilenames()  # askopenfilenames函数选择多个文件
        self.select_path.set('\n'.join(selected_files_path))  # 多个文件的路径用换行符隔开

    def select_folder(self):
        # 文件夹选择
        selected_folder = filedialog.askdirectory()  # 使用askdirectory函数选择文件夹
        self.select_path.set(selected_folder)


interface = main_interface(root)
root.mainloop()
