from selenium import webdriver
from selenium.webdriver.common.by import By
from openpyxl import Workbook,load_workbook  
from random import randint
from time import sleep

driver=webdriver.Chrome()

print("Open Website")
driver.get('https://www.w3schools.com/sql/trysql.asp?filename=trysql_select_groupby')
driver.maximize_window()

print("Find Elements")
driver.find_element(By.XPATH,'/html/body/div[2]/div/div[1]/div[1]/button').click()
sleep(3)
driver.find_element(By.XPATH,'//*[@id="yourDB"]/table/tbody/tr[2]/td[1]').click()

noOfHRows=len(driver.find_elements(By.XPATH,'//*[@id="divResultSQL"]/div/table/tbody/tr[1]'))
noOfHColumn=len(driver.find_elements(By.XPATH,'//*[@id="divResultSQL"]/div/table/tbody/tr[1]/th'))
noOfRows=len(driver.find_elements(By.XPATH,'//*[@id="divResultSQL"]/div/table/tbody/tr'))
noOfColumn=len(driver.find_elements(By.XPATH,'//*[@id="divResultSQL"]/div/table/tbody/tr[2]/td'))

print("Open New Workbook")
wb = Workbook()
ws=wb.active
ws.title="Customers"

print("Write Data In Workbook")
for r in range(1,noOfHRows+1):
    ls=[]
    for c in range(1,noOfHColumn+1):
        data = driver.find_element(By.XPATH,f"//*[@id='divResultSQL']/div/table/tbody/tr[{str(r)}]/th[{str(c)}]").text
        ls.append(data)
    ws.append(ls)
for r in range(2,noOfRows+1):
    ls=[]
    for c in range(1,noOfColumn+1):
        data = driver.find_element(By.XPATH,f"//*[@id='divResultSQL']/div/table/tbody/tr[{str(r)}]/td[{str(c)}]").text
        ls.append(data)
    ws.append(ls)

print("Save File In Current Location")
wb.save("w3school.xlsx")