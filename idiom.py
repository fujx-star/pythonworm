import bs4 as bs4
import requests
import xlwt
import xlrd
import xlutils
from xlutils.copy import copy

url = "https://www.chengyucidian.net/letter/"
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:101.0) Gecko/20100101 Firefox/101.0'
}

idioms = []


def find_idiom(href):
    res = requests.get(url="https://www.chengyucidian.net" + href, headers=headers)
    content0 = bs4.BeautifulSoup(res.content.decode("utf-8"), "lxml")
    current_idiom = content0.find(class_="page-header").text
    con = content0.find(class_="con")
    ps = con.find_all('p')
    idiom = [current_idiom, ps[0].text, ps[1].text, ps[2].text, ps[3].text]
    idioms.append(idiom)
    print(current_idiom, len(idioms))


file_path = "C:/Users/Administrator/Desktop/idiom2.xls"
book = xlrd.open_workbook(file_path)


current_url = url + str(26)
response = requests.get(url=current_url + "/p/1", headers=headers)
content = bs4.BeautifulSoup(response.content.decode("utf-8"), "lxml")
pages = content.find(class_="page")
length = len(pages.find_all("li"))
result = content.find(class_="cate").find_all('a')
for a in result:
    find_idiom(a.get('href'))

print(idioms)

sheet = book.sheet_by_index(0)
rowCount = sheet.nrows
new_book = copy(book)
new_sheet = new_book.get_sheet(0)


for i in range(len(idioms)):
    data = idioms[i]
    for j in range(0, 5):
        new_sheet.write(rowCount+i, j, data[j])
new_book.save(r"C:/Users/Administrator/Desktop/idiom2.xls")
