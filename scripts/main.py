import os
from selenium import webdriver
import pandas as pd
import csv
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

count = 0

BASE_URL = 'https://www.ozon.ru/product/'
END_URL = '?oos_search=false'

options = webdriver.ChromeOptions()
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_argument("--incognito")
options.add_argument('--start-maximized')
options.add_argument('--disable-extensions')

class DataModel:
    def __init__(
            self, full_url="", 
            product_name="", description="", 
            product_price="", rating="", 
            shop_name="", brand_name="", 
            num_of_reviews="", characteristics=""
        ):
        self.url = full_url
        self.name = product_name
        self.desc = description
        self.price = product_price
        self.rating = rating
        self.shop = shop_name
        self.brand = brand_name
        self.numrev = num_of_reviews
        self.chars = characteristics

    def __str__(self):
        return f'Name: {self.name}\nPrice: {self.price}\nRating: {self.rating}\n\n'

    def put_in_csv(self):
        pass


def parse(chrome):
    model = DataModel()
    try:
        model.name = WebDriverWait(chrome, 1).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    '//*[@id="layoutPage"]/div[1]/div[3]/div[2]/div/div/div[1]/div/h1'
                )
            )).text
        model.price = WebDriverWait(chrome, 1).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    '/html/body/div[1]/div/div[1]/div[3]/div[3] \
                    /div[2]/div[2]/div/div/div/div[1]/div/div/div/div/span/span'
                )
            )).text
        try:
            model.rating = WebDriverWait(chrome, 1).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        '//*[@id="layoutPage"]/div[1]/div[3]/div[2] \
                        /div/div/div[2]/div/div[1]/div[1]/div/div/div[1]/div/div[2]'
                    )
                )).get_attribute("style").split()[1].replace(';', '')
        except:
            model.rating = "No rating"
    except Exception as e:
        print(e)
        return 1
    print(model)
    return 0

def excel_to_csv(file):
    data_xls = pd.read_excel(file)
    data_csv = file.replace('xlsx', 'csv')
    data_xls.to_csv(data_csv, index=False, header=True)
    return data_csv

def open_window(row):
    driver = webdriver.Chrome(os.environ.get("CHROME_PATH"), options=options)
    driver.maximize_window()
    driver.get(BASE_URL + row['ID'] + END_URL)
    if parse(driver):
        open_window(row)
    driver.close()


def main():
    global count
    file = excel_to_csv('src/products.xlsx')

    with open(file) as csv_file:
        reader = csv.DictReader(csv_file)

        for row in reader:
            count += 1
            open_window(row)
            print(count)

    print(count)
    return 0


if __name__ == '__main__':
    main()