#-*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver import ActionChains
import time
import xlrd
import tkinter as tk
from  tkinter import filedialog

class ShuiKong:
    def __init__(self):
        global driver
        self.action_chains = ActionChains(driver)
    def work(self,content):
        driver.find_element_by_xpath('//a[@href="/SKServer/zzp_spbm/init.do?target=navTab&rel=zzp_zsfpkj_nav"]').click()#专票
        #href="/SKServer/zzspp_spbm/init.do?target=navTab&rel=zzspp_zsfpkj_nav"普票
        driver.find_element_by_xpath('//button[@id="xz"]').click()#专票新增按钮
        input_window = driver.find_element_by_xpath('//input[@id="spmc_1"]')#选择货物或应税劳务名称
        self.action_chains.double_click(input_window).perform()
        time.sleep(1)
        commodity_name  = content[0]
        input_window2 = driver.find_element_by_xpath('//tbody/tr[@target="slt_objId"]//td/div[contains(text(),"{}")]'.format(commodity_name))#根据模板选择第一页对应商品名称(暂时)
        self.action_chains.double_click(input_window2).perform()
        name_input = driver.find_element_by_id('ghdwmc')
        id_input = driver.find_element_by_id('ghdwdm')
        address_input = driver.find_element_by_id('ghdwdzdh')
        bank_input = driver.find_element_by_id('ghdwyhzh')
        unit_pricr_input = driver.find_element_by_id('spdj_1')
        amount_input = driver.find_element_by_id('je_1')
        remark_input = driver.find_element_by_id('bz')
        name_input.send_keys(content[1])
        id_input.send_keys(content[2])
        address_input.send_keys(content[3])
        bank_input.send_keys(content[4])
        unit_pricr_input.send_keys(content[5])
        amount_input.send_keys(content[6])
        remark_input.send_keys(content[7])
        driver.find_element_by_id('fhr').clear()
        driver.find_element_by_id('fhr').send_keys('{}'.format(str(content[8])))#复核人
        time.sleep(1)
        driver.find_element_by_id('kj').click()
        time.sleep(2)

class Excel():
    def __init__(self):
        global filename
        excel = xlrd.open_workbook(filename)
        self.table = excel.sheet_by_index(0)#索引为0的表（第0个表）
    def read(self,row):
        content = self.table.row_values(row)
        commodity_name = content[1]
        name = content[2]
        id = content[3]
        address = content[4]
        bank = content[5]
        unit_price = str(content[6])
        amount = str(content[7])
        remark = content[8]
        checker = content[9]
        return commodity_name,name,id,address,bank,unit_price,amount,remark,checker

if __name__=='__main__':
    root = tk.Tk()
    root.withdraw()
    filename = filedialog.askopenfilename()#文件对话框
    driver = webdriver.Ie()
    driver.get('http://192.168.99.181:8080/SKServer/index.jsp?relogin=true')
    time.sleep(20)
    driver.maximize_window()
    excel = xlrd.open_workbook(filename)
    table = excel.sheet_by_index(0)
    for row in range(1, table.nrows):#table.nrows表的总行数
        try:
            content = Excel().read(row)
            ShuiKong().work(content)
        except Exception as e:
            print(e)
            break