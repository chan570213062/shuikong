#-*- coding: utf-8 -*-
from tkinter import ttk,filedialog
import tkinter as tk
from selenium import webdriver
from tkinter import messagebox
from tkinter.scrolledtext import ScrolledText
import xlrd
from print_script_normal import Print_for_normal,Excel_for_print_normal
from print_script_pro import Print_fro_pro,Excel_for_print_pro
from work_script_normal import Work_for_normal,Excel_for_work_normal
from work_script_pro import Work_for_pro,Excel_for_work_pro
import time
from set_valid_date import set
import config

class Login_window(tk.Tk):
    def __init__(self):
        super(Login_window,self).__init__()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w = int((sw - 300) / 2)
        h = int((sh - 250) / 2)
        self.title('税控批量控件登录')
        self.geometry('300x250+{}+{}'.format(w, h))
        self.resizable(width=False, height=False)
        self.username = tk.StringVar()
        self.password = tk.StringVar()
        style = ttk.Style()
        style.configure('label.TLabel', font=('微软雅黑', 12))
        style.configure('button.TButton', font=('微软雅黑', 12), width=4)
        ttk.Label(self,text='用  户:',style='label.TLabel').place(x=60,y=40)
        ttk.Label(self, text='密  码:', style='label.TLabel').place(x=60, y=90)
        ttk.Entry(self, font=('宋体', 12), width=13,textvariable = self.username).place(x=120, y=42)
        ttk.Entry(self, font=('宋体', 12), width=13,textvariable = self.password,show='*').place(x=120, y=92)
        ttk.Button(self,text='确定',style='button.TButton',command = self.ConfirmLogin).place(x=65,y=150)
        ttk.Button(self,text='退出',style='button.TButton',command = self.destroy).place(x=185,y=150)

    def ConfirmLogin(self):
        username = self.username.get()
        password = self.password.get()
        if str(username) == str(config.username) and str(password) == str(config.password):
            self.destroy()
            isvalid = set()
            if isvalid==True:
                window = create_window()
                window.createWidget()
                window.mainloop()
            else:
                messagebox.showwarning('提示','已超过有效使用期限')
        else:
            messagebox.showwarning('提示','用户密码错误')

class create_window(tk.Tk):
    def __init__(self):
        super(create_window,self).__init__()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w = int((sw-650)/2)
        h = int((sh-350)/2)
        self.pro = tk.IntVar()
        self.normal = tk.IntVar()
        menubar = tk.Menu(self)
        self.filemenu = tk.Menu(menubar,tearoff=0)
        menubar.add_cascade(label='文档',menu = self.filemenu)
        self.filemenu.add_command(label='发票开具',comman = self.filedialog_work,font = ('微软雅黑',10))
        self.filemenu.add_command(label='发票打印',comman = self.filedialog_print,font = ('微软雅黑',10))
        self.config(menu = menubar)
        self.title('税控批量控件')
        self.geometry('650x350+{}+{}'.format(w,h))
        self.resizable(width=False,height=False)
        self.switch_for_work = False
        self.switch_for_print = False

    def createWidget(self):
        self.start_row = tk.IntVar()
        self.anum = tk.IntVar()
        style = ttk.Style()
        style.configure('label.TLabel',font=('微软雅黑',11))
        style.configure('button.TButton',font=('微软雅黑', 11),width=4)
        self.checkbutton1 = tk.Checkbutton(self,text='专票',font=('微软雅黑', 11),width=4,state='normal',command=self.Change_Checkbutton1,variable = self.pro)
        self.checkbutton1.place(x=30,y=15)
        self.checkbutton2 = tk.Checkbutton(self,text='普票',font=('微软雅黑', 11),width=4,state='normal',command=self.Change_Checkbutton2,variable = self.normal)
        self.checkbutton2.place(x=130,y=15)
        ttk.Label(self,text='起始行:',style='label.TLabel').place(x=70,y=65)
        ttk.Entry(self,font=('微软雅黑',11),width=7,textvariable=self.start_row).place(x=138,y=65)
        self.lable_num = ttk.Label(self,text='连续打印次数:',style='label.TLabel')
        self.lable_num.place(x=26,y=115)
        ttk.Entry(self,font=('微软雅黑',11),width=7,textvariable=self.anum).place(x=138,y=115)
        self.button_change = ttk.Button(self,text='打印',style='button.TButton',command = self.Work_and_print)
        self.button_change.place(x=30,y=165)
        ttk.Button(self,text='退出',style='button.TButton',command=self.destroy).place(x=150,y=165)
        self.src = ScrolledText(self,width=50,height=16,font=('微软雅黑', 10))
        self.src.place(x=230,y=15)

    def filedialog_work(self):
        if self.pro.get() or self.normal.get() == 1:
            root = tk.Tk()
            root.withdraw()
            self.filename = filedialog.askopenfilename()#文件对话框
            self.lable_num.config(text='连续开具次数:')
            self.button_change.config(text='开具')
            self.start_row.set('0')
            self.anum.set('0')
            self.src.insert('end','目标文件目录:{}\n'.format(self.filename))
            self.switch_for_work = True
            self.switch_for_print = False
            if not self.filename == '':
                self.driver = webdriver.Ie()
                self.driver.get('http://192.168.99.181:8080/SKServer/index.jsp?relogin=true')
                self.driver.maximize_window()
            else:
                messagebox.showwarning('提示','请选择正确有效文档')
                return self
        else:
            messagebox.showwarning('提示','请先选择 \'专票\' 或者 \'普票\'')
            return self

    def filedialog_print(self):
        if self.pro.get() or self.normal.get() == 1:
            root = tk.Tk()
            root.withdraw()
            self.filename = filedialog.askopenfilename()#文件对话框
            self.lable_num.config(text='连续打印次数:')
            self.button_change.config(text='打印')
            self.start_row.set('0')
            self.anum.set('0')
            self.src.insert('end','目标文件目录:{}\n'.format(self.filename))
            self.switch_for_print = True
            self.switch_for_work = False
            if not self.filename == '':
                self.driver = webdriver.Ie()
                self.driver.get('http://192.168.99.181:8080/SKServer/index.jsp?relogin=true')
                self.driver.maximize_window()
            else:
                messagebox.showinfo('提示','请选择正确有效文档')
                return self
        else:
            messagebox.showwarning('提示', '请先选择 \'专票\' 或者 \'普票\'')
            return self

    def Change_Checkbutton1(self):
        if self.pro.get()==1:#专票选中状态
            self.checkbutton2.deselect()
            self.filename = ''
            self.src.insert('end','切换中... \n现在选择的是[专票]\n---专票---\n')

    def Change_Checkbutton2(self):
        if self.normal.get()==1:#普票选中状态
            self.checkbutton1.deselect()
            self.filename = ''
            self.src.insert('end', '切换中... \n现在选择的是[普票]\n---普票---\n')

    def Work_and_print(self):
        start_row = self.start_row.get()
        anum = self.anum.get()
        self.total_rows = start_row
        try:
            excel = xlrd.open_workbook(self.filename)
            table = excel.sheet_by_index(0)
            valid_rows = start_row + anum
            if self.switch_for_print == True and self.pro.get() == 1 and start_row >= 2:
                confirm_message = Excel_for_print_pro(self.filename).read(start_row-1)
                self.confirm = messagebox.askokcancel('提示','正在开始打印的发票号码是\'{}\',点击\'确定\'开始打印'.format(confirm_message))
            elif self.switch_for_print == True and self.normal.get() == 1 and start_row >=2:
                confirm_message = Excel_for_print_pro(self.filename).read(start_row - 1)
                self.confirm = messagebox.askokcancel('提示', '正在开始打印的发票号码是\'{}\',点击\'确定\'开始打印'.format(confirm_message))
            if not self.total_rows > table.nrows:
                for row in range(start_row, valid_rows):
                    if self.switch_for_work == True and self.pro.get()==1:
                        try:
                            content = Excel_for_work_pro(self.filename).read(row-1)
                            # Work_for_pro(self.driver).work(content,self.driver)
                            print(content)
                            self.total_rows = self.total_rows + 1
                            self.start_row.set(valid_rows)
                            self.src.insert('end','[专票] 第{}行 已开具 {}\n'.format(row,time.strftime('%Y-%m-%d %H:%M:%S',time.localtime())))
                        except Exception as e:
                            print(e)
                            messagebox.showinfo('提示', '已超出文档限制,检查是否已全部开具完毕')
                            self.start_row.set('0')
                            self.anum.set('0')
                            self.filename = ''
                            break
                    elif self.switch_for_print == True and self.pro.get()==1 and self.confirm == True:
                        try:
                            content = Excel_for_print_pro(self.filename).read(row-1)
                            # Print_fro_pro(self.driver).work(content,self.driver)
                            print(content)
                            self.total_rows = self.total_rows + 1
                            self.start_row.set(valid_rows)
                            self.src.insert('end','[专票] {} 打印完毕{}\n'.format(content,time.strftime('%Y-%m-%d %H:%M:%S',time.localtime())))
                        except Exception as e:
                            print(e)
                            messagebox.showinfo('提示', '已超出文档限制,检查是否已全部打印完毕')
                            self.start_row.set('0')
                            self.anum.set('0')
                            self.filename = ''
                            break
                    elif self.switch_for_work ==True and self.normal.get()==1:
                        try:
                            content = Excel_for_work_normal(self.filename).read(row-1)
                            # Work_for_normal(self.driver).work(content,self.driver)
                            print(content)
                            self.total_rows = self.total_rows + 1
                            self.start_row.set(valid_rows)
                            self.src.insert('end','[普票] 第{}行 已开具 {}\n'.format(row,time.strftime('%Y-%m-%d %H:%M:%S',time.localtime())))
                        except Exception as e:
                            print(e)
                            messagebox.showinfo('提示', '已超出文档限制,检查是否已全部开具完毕')
                            self.start_row.set('0')
                            self.anum.set('0')
                            self.filename = ''
                            break
                    elif self.switch_for_print == True and self.normal.get() == 1 and self.confirm == True:
                        try:
                            content = Excel_for_print_normal(self.filename).read(row-1)
                            # Print_for_normal(self.driver).work(content,self.driver)
                            print(content)
                            self.total_rows = self.total_rows + 1
                            self.start_row.set(valid_rows)
                            self.src.insert('end','[普票] {} 打印完毕{}\n'.format(content,time.strftime('%Y-%m-%d %H:%M:%S',time.localtime())))
                        except Exception as e:
                            print(e)
                            messagebox.showinfo('提示', '已超出文档限制,检查是否已全部打印完毕')
                            self.start_row.set('0')
                            self.anum.set('0')
                            self.filename = ''
                            break
                else:
                    if self.total_rows > table.nrows:
                        messagebox.showinfo('提示', '已超出文档限制,检查是否已全部打印完毕')
                        self.start_row.set('0')
                        self.anum.set('0')
                        self.filename = ''
                        return self
        except Exception as e:
            print(e)
            messagebox.showwarning('提示','请选择文档')
            return self

if __name__=='__main__':
    loginwindow = Login_window()
    loginwindow.mainloop()
    # isvalid = set()
    # if isvalid==True:
    #     window = create_window()
    #     window.createWidget()
    #     window.mainloop()
    # else:
    #     messagebox.showwarning('提示','已超过有效使用期限')
