import urllib.request as ul
from bs4 import BeautifulSoup as soup

BASE_URL = 'http://www.ozon.ru/product/'
END_URL = '?oos_search=false'

req = ul.Request(BASE_URL + '269997508' + END_URL, headers={'User-Agent': 'Mozilla/5.0'})
client = ul.urlopen(req)
htmldata = client.read()
client.close()

pagesoup = soup(htmldata, "html.parser")
print(pagesoup)

