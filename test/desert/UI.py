from tkinter import *
from tkinter.ttk import *
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog
import tkinter as tk

"""
全局通用函数
"""


# 自动隐藏滚动条
def scrollbar_autohide(bar, widget):
    def show():
        bar.lift(widget)

    def hide():
        bar.lower(widget)

    hide()
    widget.bind("<Enter>", lambda e: show())
    bar.bind("<Enter>", lambda e: show())
    widget.bind("<Leave>", lambda e: hide())
    bar.bind("<Leave>", lambda e: hide())


# 窗口类，继承于Tk
class WinGUI(ttk.Window):  # 让这个窗口继承于ttk的Window
    def __init__(self):
        super().__init__()
        self.__win()  # 初始化时调用自身的函数，就可以创建了一个窗口
        self.frame = Frame(self)  # 实例化容器类的一个容器，这个容器马上就有了各种组件

    def __win(self):
        self.title("报表整理工具")
        self.themename = 'vapor'
        # 设置窗口大小、居中
        width = 1600
        height = 1200
        screenwidth = self.winfo_screenwidth()
        screenheight = self.winfo_screenheight()
        geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.geometry(geometry)
        self.resizable(width=False, height=False)


# 容器类，继承于Frame，本类实现了许多组件的自动挂载
class Frame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.__frame()  # 调用自身的方法，把自身放置到某个坐标位置。
        label = Label(self, text="报表整合工具", anchor="center")
        label.place(x=10, y=10, width=219, height=50)
        label = Label(self, text="更换皮肤", anchor="center")
        label.place(x=1350, y=10, width=100, height=36)
        select_skin = self.__tk_select_box_sele_skin()
        select_skin.place(x=1450, y=10, width=100, height=36)
        # 工单
        label = Label(self, text="选择工单报表（多选）", anchor="center")
        label.place(x=10, y=80, width=200, height=36)
        self.select_path_WO = ttk.StringVar()
        ipt = Entry(self, textvariable=self.select_path_WO)
        ipt.place(x=10, y=110, width=400, height=36)
        btn = Button(self, text="选择文件", command=self.select_file_WO)
        btn.place(x=230, y=70, width=100, height=36)
        # 未关闭报表
        label = Label(self, text="未关闭Bug报表", anchor="center")
        label.place(x=10, y=160, width=150, height=36)
        self.select_path_Un = ttk.StringVar()
        ipt = Entry(self, textvariable=self.select_path_Un)
        ipt.place(x=10, y=190, width=400, height=36)
        btn = Button(self, text="选择文件", command=self.select_file_Un)
        btn.place(x=230, y=150, width=100, height=36)
        # 所有报表
        label = Label(self, text="所有Bug报表", anchor="center")
        label.place(x=10, y=240, width=130, height=36)
        self.select_path_All = ttk.StringVar()
        ipt = Entry(self, textvariable=self.select_path_All)
        ipt.place(x=10, y=270, width=200, height=36)
        btn = Button(self, text="选择文件", command=self.select_file_All)
        btn.place(x=230, y=270, width=100, height=36)
        # 目标文件
        label = Label(self, text="目标文件", anchor="center")
        label.place(x=10, y=330, width=100, height=36)
        self.select_path_Tag = ttk.StringVar()
        ipt = Entry(self, textvariable=self.select_path_Tag)
        ipt.place(x=10, y=360, width=200, height=36)
        btn = Button(self, text="选择文件", command=self.select_file_Tag)
        btn.place(x=230, y=360, width=100, height=36)
        # 启动按钮
        btn = Button(self, text="合并")
        btn.place(x=100, y=500, width=138, height=51)
        # 输出框
        text = Text(self, borderwidth=5, bg='#B0E0E6')
        text.place(x=480, y=70, width=1100, height=1100)

    def __frame(self):  # 调用自己的方法，把自己放置到某个位置
        self.place(x=0, y=0, width=1600, height=1200)

    def __tk_select_box_sele_skin(self):
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
    def select_file_WO(self):
        # 多个文件选择
        selected_files_path = filedialog.askopenfilenames()  # askopenfilenames函数选择多个文件
        self.select_path_WO.set('\n'.join(selected_files_path))  # 多个文件的路径用换行符隔开

    def select_file_Un(self):
        # 单个文件选择
        selected_file_path = filedialog.askopenfilename()  # 使用askopenfilename函数选择单个文件
        self.select_path_Un.set(selected_file_path)

    def select_file_All(self):
        # 单个文件选择
        selected_file_path = filedialog.askopenfilename()  # 使用askopenfilename函数选择单个文件
        self.select_path_All.set(selected_file_path)

    def select_file_Tag(self):
        # 单个文件选择
        selected_file_path = filedialog.askopenfilename()  # 使用askopenfilename函数选择单个文件
        self.select_path_Tag.set(selected_file_path)

    # def select_folder(self):
    #     # 文件夹选择
    #     selected_folder = filedialog.askdirectory()  # 使用askdirectory函数选择文件夹
    #     self.select_path.set(selected_folder)


class Win(WinGUI):
    def __init__(self):
        super().__init__()
        self.__event_bind()

    def __event_bind(self):
        pass


if __name__ == "__main__":
    win = Win()
    win.mainloop()
