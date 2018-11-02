#-*- coding: utf-8 -*-
from selenium import webdriver
from  selenium.webdriver import ActionChains
import xlrd
import time
import tkinter as tk
from tkinter import filedialog

class Print_for_normal():
    def __init__(self,driver):
        self.action_chains = ActionChains(driver)

    def work(self,content,driver):
        try:
            driver.find_element_by_xpath('//td/input[@name="fphm"]').clear()
        except Exception as e:
            input_check = driver.find_element_by_xpath('//*[@id="sidebar"]//a[@href="/SKServer/zzspp/fpcx.do?target=navTab&rel=zzspp_fpcx_nav"]')#普票查询,会变化
            self.action_chains.double_click(input_check).perform()
            fpnum_input = driver.find_element_by_xpath('//td/input[@name="fphm"]')
            fpnum_input.send_keys(content)
            driver.find_element_by_id('cx').click()
            driver.find_element_by_xpath('//div[@class="gridTbody"]//div[contains(text(),{})]'.format(content)).click()
            driver.find_element_by_id('dy').click()
            # time.sleep(10)
        else:
            fpnum_input = driver.find_element_by_xpath('//td/input[@name="fphm"]')
            fpnum_input.send_keys(content)
            driver.find_element_by_id('cx').click()
            driver.find_element_by_xpath('//div[@class="gridTbody"]//div[contains(text(),{})]'.format(content)).click()
            driver.find_element_by_id('dy').click()
            # time.sleep(10)

class Excel_for_print_normal():
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
    time.sleep(30)
    driver.maximize_window()
    excel = xlrd.open_workbook(filename)
    table = excel.sheet_by_index(0)
    for row in range(1,table.nrows):
        try:
            content = Excel_for_print_normal(filename).read(row)
            Print_for_normal(driver).work(content,driver)
        except Exception as e:
            print(e)
            break