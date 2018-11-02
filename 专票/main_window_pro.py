#-*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from tkinter import filedialog
from tkinter import messagebox
import xlrd
from print_script_pro_with_window import Print,Excel

class create_window(tk.Tk):
    def __init__(self):
        super(create_window,self).__init__()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w = int((sw-270)/2)
        h = int((sh-200)/2)
        menubar = tk.Menu(self)
        self.filemenu = tk.Menu(menubar,tearoff=0)
        menubar.add_cascade(label='文档',menu = self.filemenu)
        self.filemenu.add_command(label='打开',comman = self.filedialog )
        self.config(menu = menubar)
        self.title('税控打印(专票)')
        self.geometry('270x200+{}+{}'.format(w,h))

    def createWidget(self):
        self.start_row = tk.IntVar()
        self.anum = tk.IntVar()
        style = ttk.Style()
        style.configure('label.TLabel',font=('新宋体',13))
        style.configure('button.TButton',font=('新宋体', 13),width=4)
        ttk.Label(self,text='起始行:',style='label.TLabel').place(x=82,y=30)
        ttk.Entry(self,font=('新宋体',13),width=10,textvariable=self.start_row).place(x=150,y=30)
        ttk.Label(self,text='连续打印次数:',style='label.TLabel').place(x=30,y=80)
        ttk.Entry(self,font=('新宋体',13),width=10,textvariable=self.anum).place(x=150,y=80)
        ttk.Button(self,text='打印',style='button.TButton',command = self.start).place(x=60,y=130)
        ttk.Button(self,text='退出',style='button.TButton',command = window.quit).place(x=170,y=130)

    def filedialog(self):
        root = tk.Tk()
        root.withdraw()
        self.filename = filedialog.askopenfilename()#文件对话框
        if not self.filename == '':
            self.driver = webdriver.Ie()
            self.driver.get('http://192.168.99.181:8080/SKServer/index.jsp?relogin=true')
            self.driver.maximize_window()
        else:
            messagebox.showinfo('提示','请选择正确有效文档')
            return window

    def start(self):
        start_row = self.start_row.get()
        anum = self.anum.get()
        self.total_rows = start_row
        try:
            excel = xlrd.open_workbook(self.filename)
            table = excel.sheet_by_index(0)
            valid_rows = start_row + anum
            if not self.total_rows > table.nrows:
                for row in range(start_row,valid_rows):
                    try:
                        content = Excel(self.filename).read(row-1)
                        print(content)
                        Print(self.driver).work(content,self.driver)
                        self.total_rows = self.total_rows + 1
                    except Exception as e:
                        print(e)
                        messagebox.showinfo('提示','已超出文档限制,检查是否已全部打印完毕')
                        self.start_row.set('0')
                        self.anum.set('0')
                        break
                else:
                    self.start_row.set(valid_rows)
                    if self.total_rows > table.nrows:
                        messagebox.showinfo('提示','已超出文档限制,检查是否已全部打印完毕')
                        self.start_row.set('0')
                        self.anum.set('0')
            else:
                messagebox.showinfo('提示', '已超出文档限制,检查是否已全部打印完毕')
                self.start_row.set('0')
                self.anum.set('0')
        except Exception as e:
            print(e)
            messagebox.showwarning('提示','请选择文档')
            return window

if __name__=='__main__':
    window = create_window()
    window.createWidget()
    window.mainloop()