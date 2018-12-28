#-*- coding: utf-8 -*-
from selenium import webdriver
from  selenium.webdriver import ActionChains
import xlrd
import time
import tkinter as tk
from tkinter import filedialog

class Zuofei_for_pro():
    def __init__(self,driver):
        self.action_chains = ActionChains(driver)

    def work(self,content,driver):
        try:
            driver.find_element_by_xpath('//td/input[@name="fphm"]').clear()
        except Exception as e:
            input_check = driver.find_element_by_xpath('//a[@href="/SKServer/zzp/fpcx.do?target=navTab&rel=zzp_fpcx_nav"]').click()#专票查询
            self.action_chains.double_click(input_check).perform()
            time.sleep(1)
            fpnum_input = driver.find_element_by_xpath('//td/input[@name="fphm"]')
            fpnum_input.send_keys(int(content))
            time.sleep(0.5)
            driver.find_element_by_id('cx').click()
            time.sleep(0.5)
            driver.find_element_by_xpath('//div[@class="gridTbody"]//div[contains(text(),{})]'.format(int(content))).click()
            driver.find_element_by_id('zf').click()
            time.sleep(0.5)
            driver.find_element_by_xpath('//div[@id="alertMsgBox"]//a[1]').click()
            time.sleep(4.5)
        else:
            fpnum_input = driver.find_element_by_xpath('//td/input[@name="fphm"]')
            fpnum_input.send_keys(int(content))
            time.sleep(0.5)
            driver.find_element_by_id('cx').click()
            time.sleep(0.5)
            driver.find_element_by_xpath('//div[@class="gridTbody"]//div[contains(text(),{})]'.format(int(content))).click()
            driver.find_element_by_id('zf').click()
            time.sleep(0.5)
            driver.find_element_by_xpath('//div[@id="alertMsgBox"]//a[1]').click()
            time.sleep(4.5)

class Excel_for_print_pro():
    def __init__(self,filename):
        excel = xlrd.open_workbook(filename)
        self.table = excel.sheet_by_index(0)

    def read(self,row):
        content = self.table.row_values(row)
        fpnum = content[2]
        return fpnum

if __name__=='__main__':
    root = tk.Tk()
    root.withdraw()
    filename = filedialog.askopenfilename()
    driver = webdriver.Ie()
    driver.get('http://192.168.99.181:8080/SKServer/index.jsp?relogin=true')
    time.sleep(15)
    driver.maximize_window()
    excel = xlrd.open_workbook(filename)
    table = excel.sheet_by_index(0)
    for row in range(1,table.nrows):
        try:
            content = Excel_for_print_pro(filename).read(row)
            Zuofei_for_pro(driver).work(content,driver)
        except Exception as e:
            print(e)
            break