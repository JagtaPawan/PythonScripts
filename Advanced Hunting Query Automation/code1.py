#Neccessary Imports
from cmath import exp
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

url= "https://security.microsoft.com/v2/advanced-hunting?query=H4sIAAAAAAAAA71XbW_SUBQ-n038D82-bDNogM0lMyERGSYk2ySO6AcxppTCOmmL0L1o_PE-5_QWWnpLW4oLgd6ee16ee97u4YJseiCHLDwHWLl0g5VF97TAs0M-eTTB_lQoJgVYM61NS3xs-XXx9LDzkl7QXzLokW5BYXlji4YejbHvYGXQER1Cj4W3Mb2mJtWpge8h1TT0Rgb9NIPezKCfZNDPtPS64jcSO-eC8nh1cpuecD72Bp-tB842fNrFu0kjmmFnTC3IOvDIUa5vWvgYdKDxzAHoZswKa3JpDguO0AORbQje8LcOlDXByfyfgcQEN9vj6PVh3Rfk1grDvlE3K6Nu0xdkJ9v1xGqYo5ey5kwM9oCysReU7N855BZi1QOG6shOKyPrw56597ieVMbVAacPm-NnycOzyngH4GJ-7nH_H3F9Dx7ugtsE7wz1Y4mEp_Jyf1jPd-5LUfdcwpIrSBfg_yP3x0XsfroWv9vQwLp-A6VuN1mL6e7L-p8gq9-v2iXj2svLl-12ybPk8RfpUpsa0zz5HSWuQ8exe-3HNZeRrFK9cZvlZHepQEPqKm6zqGRYRyPURbJquF650gYqCzkTwlMw9RO09qGBs2QicXZX9XiH93A--ylP9sYEHp8pzwfiB57eWCJu86PgtYHwQc2GS-102JOzBNIdovwK_Wqp2TLsIK7qKpfCw5KW-DsQz3jg4150hTOFFXAj-T-iNzKR2djTT2jXWD0iyxnPUsWwhX1fvLRQiI7wFvawQGKyG-aaYOTvMX3Dbf49p_Pxeko_Vrmgw8o6X0GPkRH354tkxBd14fUNcQUJjkxXPJWOSFKLDy1jVV_s31ttrN9DvkPvaKg8P5UbygVyU_iGsMj4FxJJH2cNQPuqzu7D5lJ5a6JygS0OU5UwzMycTc5yOaQ7Z5gfQ7HJ67eI6zpD5nKWO9Vt9BW-HRNr4kzxJQ-yMmXTjr5rbEobQs3K0HxcRWtAjyavAvJPrasPRzotI72Xiv-Fp62tkZ5MSr42ox1g9dQ9wXgDmVu4636APkeyoFwfKO-DqM-nI5PEUZP_65yjYaXr41QkE3XWytoqn6vFbz3d3FijrH_s0U75eW6tM282i3OmZ65oVzdNRXtl5qFIptw8E0kVnUj0_m8U6AnbuoluYjFUflXL8H_HPktrDBMAAA&timeRangeId=" #day

driver.get(url)
time.sleep(5)
user_name = driver.find_element_by_xpath('//*[@id="i0116"]')
user_name.send_keys(USERNAME)

submit_button = driver.find_element_by_xpath('//*[@id="idSIButton9"]')
submit_button.click()

time.sleep(10)
password = driver.find_element_by_xpath('//*[@id="i0118"]')
password.send_keys(PASSWORD)

submit_button = driver.find_element_by_xpath('//*[@id="idSIButton9"]')
submit_button.click()

time.sleep(10)
Yes = driver.find_element_by_xpath('//*[@id="idSIButton9"]')
Yes.click()

time.sleep(20)
runQuery = driver.find_element_by_xpath('//*[@id="pageContentZone"]/div[4]/div/div[2]/div/div/div[2]/div/div[1]/div/div/div/div/div/div[1]/div[1]/div/button')
# //*[@id="pageContentZone"]/div/div/div[2]/div/div[2]/div/div[1]/div[2]/div/div/div[1]/div/div/div/div/div[1]/div[1]/div/button
runQuery.click()

time.sleep(10)
export = driver.find_element_by_xpath('//*[@id="pageContentZone"]/div[4]/div/div[2]/div/div/div[3]/div/div[3]/div[2]/div/div/div/div/div/div/div[1]/div/div/button')
# //*[@id="pageContentZone"]/div/div/div[2]/div/div[3]/div/div[3]/div[2]/div/div/div/div/div/div/div[1]/div/div/button
export.click()

time.sleep(20)
driver.quit()
