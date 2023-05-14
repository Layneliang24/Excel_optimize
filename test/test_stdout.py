import sys
import time
import tkinter


class mystdout(object):
    def __init__(self, text):
        self.stdoutbak = sys.stdout
        self.stderrbak = sys.stderr
        sys.stdout = self
        sys.stderr = self
        self.text = text

    def write(self, info):  # 外部的print语句将执行本write()方法，并由当前sys.stdout输出
        # print('调用了尼玛')
        # info信息即标准输出stdout/stderr接收到的输出信息。
        self.text.insert('end', info)
        self.text.update()
        self.text.see(tkinter.END)


def btn_func():
    for i in range(5):
        print(i)
        time.sleep(1)


window = tkinter.Tk()
t = tkinter.Text(window)
t.pack()
mystd = mystdout(t)  # 只要实例化了重定向类，那么函数的执行输出就换了地方了。
b = tkinter.Button(window, text='Start', command=btn_func)
b.pack()

window.mainloop()
