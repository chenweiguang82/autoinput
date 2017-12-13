# -*- coding: utf-8 -*-
"""
Created on Tue Dec 12 15:06:49 2017

@author: ChenWeiguang
"""

from selenium import webdriver
import time, sys
import xlrd

def AutoInput(driver,table,start,end):
    S01=driver.find_element_by_name('S01')
    I=[]
    for i in range(4):
        I.append(driver.find_element_by_name('I%02d' %(i+1)))
    okbtn=driver.find_element_by_id('okbtn')
    for i in range(start,end+1):
        print('填写仪器序号：%3d' %(table.row_values(i)[0]))
        S01.clear()
        S01.send_keys(table.row_values(i)[1])
        for j in range(4):
            I[j].clear()
            I[j].send_keys(str(int(table.row_values(i)[j+2])))
        time.sleep(0.5)
        okbtn.click()
        time.sleep(1)
        
def Readxlsx(filename):
    data=xlrd.open_workbook('C:/Users/ChenWeiguang/Desktop/sbqk.xlsx')
    table=data.sheet_by_index(0)
    return table

def CountDown(Seconds):
    for i in range(Seconds):    
        sys.stdout.write('\r%3d' %(Seconds-i))
        time.sleep(1)
    print('')
    
def Login(driver):
    username=driver.find_element_by_name('userName')
    username.clear()
#用户名
    username.send_keys('xxxxx')
    time.sleep(0.5)
    password=driver.find_element_by_name('password')
    password.clear()
#密码
    password.send_keys('******')
    time.sleep(0.5)
    submit=driver.find_element_by_class_name('btn')
    submit.click()

def GotoSbqk(driver):
    sbqk=driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr[3]/td/table[2]/tbody/tr[33]/td[6]/a[1]')
    sbqk.click()

driver = webdriver.Chrome('C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe')
driver.get('http://zypt.neusoft.edu.cn/hasdb/login_view.action')

table = Readxlsx('C:/Users/ChenWeiguang/Desktop/sbqk.xlsx')
Login(driver)
GotoSbqk(driver)
#填报信息
AutoInput(driver,table,1,100)

CountDown(10)
