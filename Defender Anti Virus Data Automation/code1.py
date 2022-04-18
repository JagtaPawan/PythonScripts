#Neccessary Imports

from os import path
from selenium import webdriver
import time
import pyautogui
import datetime

path = "C:\Program Files (x86)\chromedriver.exe"
driver = webdriver.Chrome(path)

x = datetime.datetime.now()

today = x.strftime("%A-%d-%m-%Y")

driver.maximize_window()

USERNAME = ""
PASSWORD = ""

url= "https://security.microsoft.com/machines"

driver.get(url)

time.sleep(10)

user_name = driver.find_element_by_xpath('//*[@id="i0116"]')
user_name.send_keys(USERNAME)

submit_button = driver.find_element_by_xpath('//*[@id="idSIButton9"]')
submit_button.click()

time.sleep(10)

password = driver.find_element_by_xpath('//*[@id="i0118"]')
password.send_keys(PASSWORD)

submit_button = driver.find_element_by_xpath('//*[@id="idSIButton9"]')
submit_button.click()

time.sleep(20)
Yes = driver.find_element_by_xpath('//*[@id="idSIButton9"]')
Yes.click()

time.sleep(20)
dropdown = driver.find_element_by_xpath('//*[@id="button_fancy-select-0"]/wcd-shared-icon[2]/icon-switch/fab-icon/i')
dropdown.click()

time.sleep(5)
oneday = driver.find_element_by_xpath('/html/body/section[3]/ul/li[1]/div')
oneday.click()

time.sleep(5)
export = driver.find_element_by_xpath('//*[@id="command-bar-item-1"]/wcd-shared-icon/icon-switch/fab-icon/i')
export.click() 

time.sleep(20)
driver.quit()
