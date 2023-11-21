#Import Packages
from pandas import ExcelWriter
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import pandas as pd
import time

#Setup WebDriver
option = webdriver.ChromeOptions()
option.add_argument("--headless")
service = Service("chromedriver.exe")
driver = webdriver.Chrome(service=service, options=option)

#Scraping Data
shopee_link = "https://shopee.co.id/search?keyword=laptop"
driver.set_window_size(1300, 800)
driver.get(shopee_link)

distance = 500
for i in range(1, 7):
    last = distance * i
    command = "window.scrollTo(0," + str(last) + ")"
    driver.execute_script(command)
    print("loading to-" + str(i))
    time.sleep(1)

time.sleep(5)
driver.save_screenshot("home.png")
content = driver.page_source
driver.quit()

data = BeautifulSoup(content, "html.parser")

#Extracting Data
i = 1
base_url = "https://shopee.co.id"

list_name, list_price, list_link, list_sold, list_location = [], [], [], [], []


for area in data.find_all("div", class_="col-xs-2-4 shopee-search-item-result__item"):
    print("processing data to-" + str(i))
    name = area.find("div", class_="ie3A+n bM+7UW Cve6sh").get_text()
    price = area.find("span", class_="ZEgDH9").get_text()
    link = base_url + area.find("a")["href"]
    sold = area.find("div", class_="r6HknA uEPGHT")
    if sold is not None:
        sold = sold.get_text()
    location = area.find("div", class_="zGGwiV").get_text()

    list_name.append(name)
    list_price.append(price)
    list_link.append(link)
    list_sold.append(sold)
    list_location.append(location)
    i += 1
    print("------")

#Storing Data
df = pd.DataFrame({"Name": list_name, "Price": list_price, "Link": list_link, "Sold": list_sold,
                   "Location": list_location})
writer: ExcelWriter = pd.ExcelWriter("Result/Laptop.xlsx")
df.to_excel(writer, "Sheet1", index=False)
writer.save()
