from selenium import webdriver
from openpyxl import Workbook
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import datetime


wb = Workbook()
ws = wb.active

date = datetime.datetime.now().strftime("%d-%m-%Y")
s = Service(r"D:\Durai\Driver\chromedriver.exe")
driver = webdriver.Chrome(service=s)


url = "https://www.amazon.in/Microsoft-i5-1035G1-Touchscreen-Graphics-THH-00023/dp/B08SX5XVBK/ref=sr_1_2?dchild=1&keywords=Microsoft+Surface+Go+Intel+Core+i5+10th+Gen+Windows+10+Home+Laptop%28Platinum%2C8GB-128GB%29&qid=1628601251&s=computers&sr=1-2"
driver.get(url)
price = driver.find_element(By.CLASS_NAME, "apexPriceToPay")

print(price.text)