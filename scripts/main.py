import os
from selenium import webdriver
import pandas as pd
import csv
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
count = 0

BASE_URL = 'https://www.ozon.ru/product/'
END_URL = '?oos_search=false'

lst = os.listdir('src/cards')
NUM_FILES = len(lst)


options = webdriver.ChromeOptions()
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_argument("--incognito")
options.add_argument('--start-maximized')
options.add_argument('--disable-extensions')

class DataModel:

    status = ""
    out_of_stock = False
    header = {}

    def __init__(
            self, full_url="", 
            product_name="", description="", 
            product_price="", rating="", 
            shop_name="", num_of_reviews="",
            characteristics={}, old_price="", product_id=""
        ):
        self.url = full_url
        self.name = product_name
        self.desc = description
        self.price = product_price
        self.rating = rating
        self.shop = shop_name
        self.numrev = num_of_reviews
        self.chars = characteristics
        self.oldprice = old_price
        self.id = product_id

    def __str__(self):
        return f'Name: {self.name}\nPrice: {self.price}\nRating: \
        {self.rating}\nNumber of reviews: {self.numrev}\nDescription: \
        {self.desc}\nCharacteristics: {self.chars}\nShop: {self.shop} \
        \n\nOut of stock: {self.out_of_stock}'
    
    def make_dict(self):
        self.header['Название товара'] = self.name
        self.header['Описание товара'] = self.desc
        self.header['Цена в рублях'] = self.price
        self.header['Цена без скидки'] = self.oldprice
        self.header['Оценка покупателей'] = self.rating
        self.header['Название магазина'] = self.shop
        self.header['Количество отзывов'] = self.numrev

        self.header = self.header | self.chars
        self.header['Адрес товара'] = self.url

    def put_in_csv(self):
        self.make_dict()
        with open(f'src/cards/{self.id}.tsv', 'w') as file:
            writer = csv.writer(file, delimiter='\t')
            writer.writerows(self.header.items())
            writer.writerow(['Статус ', self.status])
            


def parse(chrome, product_url, iden):
    model = DataModel(full_url=product_url, product_id=iden)
    try:
        temp_status = WebDriverWait(chrome, 1).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    '/html/body/div[1]/div/div[1]/div[2]/div[1]/h2'
                )
            )).text.replace(" ", "")
        model.status = "Отсутствует"
        model.put_in_csv()
        return 0
    except:
        try:
            temp_status = WebDriverWait(chrome, 1).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        '/html/body/div[1]/div/div[1]/div[3]/div[1]/div/div[3]/h2'
                    )
                )).text.replace(" ", "")
            if temp_status == "Этоттоварзакончился":
                model.out_of_stock = True
                model.status = "Закончился"
            elif temp_status == "Товарнедоставляетсяввашрегион":
                model.status = "Не доставляется"
        except:
            model.status = "В наличии"
    try:
        try: 
            model.name = WebDriverWait(chrome, 1).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        '//*[@id="layoutPage"]/div[1]/div[3]/div[2]/div/div/div[1]/div/h1'
                    )
                )).text
        except:
            model.name = WebDriverWait(chrome, 1).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        '/html/body/div[1]/div/div[1]/div[3]/div[3]/div[3]/div/div[3]/h1'
                    )
                )).text
        
        try: 
            model.price = "".join(WebDriverWait(chrome, 1).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        '/html/body/div[1]/div/div[1]/div[3]/div[3] \
                        /div[2]/div[2]/div/div/div/div[1]/div/div/div/div/span/span'
                    )
                )).text.split()[:-1])
        except:
            try:
                model.price = "".join(WebDriverWait(chrome, 1).until(
                    EC.presence_of_element_located(
                        (
                            By.XPATH,
                            '/html/body/div[1]/div/div[1]/div[3]/div[3] \
                            /div[3]/div/div[10]/div/div/div[1]/div/div/div[1]/div/span[1]/span'
                        )
                    )).text.split()[:-1])
            except:
                model.price = "-"
        
        try:
            model.oldprice = "".join(WebDriverWait(chrome, 1).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        '/html/body/div[1]/div/div[1]/div[3]/div[3]/ \
                        div[2]/div[2]/div/div/div/div[1]/div/div/div/div/span[2]'
                    )
                )).text.split()[:-1])
        except:
            try:
                model.oldprice = "".join(WebDriverWait(chrome, 1).until(
                    EC.presence_of_element_located(
                        (
                            By.XPATH,
                            '/html/body/div[1]/div/div[1]/div[3]/div[3]/ \
                            div[3]/div/div[10]/div/div/div[1]/div/div/div[1]/div/span[2]'
                        )
                    )).text.split()[:-1])
            except:
                model.oldprice = "-"
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
            try:
                model.rating = WebDriverWait(chrome, 1).until(
                    EC.presence_of_element_located(
                        (
                            By.XPATH,
                            '/html/body/div[1]/div/div[1]/div[3]/div[3]/ \
                            div[3]/div/div[5]/div[1]/div/div/div[1]/div/div[2]'
                        )
                    )).get_attribute("style").split()[1].replace(';', '')
            except:
                model.rating = "No rating"
        try: 
            model.shop = WebDriverWait(chrome, 1).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        '/html/body/div[1]/div/div[1]/div[5]/div/div[1]/ \
                        div[2]/div/div[1]/div/div[1]/div[1]/div/div[2]/div[1]'
                    )
                )).text
        except:
            try: 
                model.shop = WebDriverWait(chrome, 1).until(
                    EC.presence_of_element_located(
                        (
                            By.XPATH,
                            '/html/body/div[1]/div/div[1]/div[3]/div[3]/div[3] \
                            /div/div[24]/div/div[1]/div[1]/div/div[2]/div[1]/div/a'
                        )
                    )).text
            except:
                model.shop = WebDriverWait(chrome, 1).until(
                    EC.presence_of_element_located(
                        (
                            By.XPATH,
                            '/html/body/div[1]/div/div[1]/div[5]/div/div[1]/ \
                            div[2]/div/div[1]/div/div[1]/div[1]/div/div/div[1]/div[2]/a'
                        )
                    )).text
        
        try: 
            model.desc = " ".join(WebDriverWait(chrome,1).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    '//*[@id="section-description"]'
                )
            )).text.split("\n")[1:])
        except:
            model.desc = "No description"
        
        try: 
            temp_chars = WebDriverWait(chrome, 1).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    '//*[@id="section-characteristics"]'
                )
            )).text.split("\n")[1:-25]
            model.chars = dict(zip(temp_chars[0::2], temp_chars[1::2]))
        except:
            model.char = {'Характеристика': "-"}
        try:
            model.numrev = WebDriverWait(chrome, 1).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        '/html/body/div[1]/div/div[1]/div[5]/ \
                        div/div[1]/div[3]/div[3]/div/div[2]/div/div'
                    )
                )).text
        except:
            model.numrev = WebDriverWait(chrome, 1).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        '/html/body/div[1]/div/div[1]/div[4]/ \
                        div/div[4]/div/div[2]/div/div[2]/div/div'
                    )
                )).text

    except Exception as e:
        print(model)
        return 1
    model.put_in_csv()
    return 0

def excel_to_csv(file):
    data_xls = pd.read_excel(file)
    data_csv = file.replace('xlsx', 'csv')
    data_xls.to_csv(data_csv, index=False, header=True)
    return data_csv

def open_window(row):
    driver = webdriver.Chrome(os.environ.get("CHROME_PATH"), options=options)
    driver.maximize_window()
    product_url = BASE_URL + row['ID'] + END_URL
    driver.get(product_url)
    if parse(driver, product_url, row['ID']):
        open_window(row)
    driver.close()


def main():
    global count
    file = excel_to_csv('src/products.xlsx')

    with open(file) as csv_file:
        reader = csv.DictReader(csv_file)

        for row in reader:
            if count >= NUM_FILES:
                open_window(row)
            count += 1
            print(count)

    print(count)
    return 0


if __name__ == '__main__':
    main()