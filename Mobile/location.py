from openpyxl import Workbook
from selenium import webdriver
import datetime
import time
# date = datetime.date.today().strftime("%d-%m-%Y")
# l = 2
#
# wb = Workbook()
# ws = wb.active
# ws.title = "MOBILE"

driver = webdriver.Chrome(executable_path=r"D:\Durai\Driver\chromedriver.exe")

# ws.cell(row=1, column=1).value = datetime.datetime.today()

driver.get(url="https://www.flipkart.com/search?q=mobiles&sid=tyy%2C4io&as=on&as-show=on&otracker=AS_QueryStore_OrganicAutoSuggest_2_2_na_na_na&otracker1=AS_QueryStore_OrganicAutoSuggest_2_2_na_na_na&as-pos=2&as-type=RECENT&suggestionId=mobiles%7CMobiles&requestId=fed5203a-42e4-42d1-a23b-40aa938c45e4&as-searchtext=mo")

for name in driver.find_elements_by_class_name("B_NuCI"):
    print(name.text)
