from urllib.request import urlopen
from bs4 import BeautifulSoup
import xlwt

def Yating_deng():
    workbook=xlwt.Workbook(encoding='utf-8')
    booksheet=workbook.add_sheet('YSL', cell_overwrite_ok=True)
    booksheet.write(0,0,'Product')
    booksheet.write(0,1,'Price')
    row = 1
    
    html = urlopen("https://www.nocibe.fr/yves-saint-laurent/C-47287/?allProduct=true")
    bsObj = BeautifulSoup(html, "html.parser")
    try:
        bsObj_list = bsObj.find("div", "container-rwd under-fixed is-banner").find_all("article",limit=50)
    except AttributeError:
        print("页面缺少一些属性！不过不用担心！")

    for article in bsObj_list:
        product = article.find("a",class_="brown one-product")
        if (article.find("span","price-regular") == None):
            prix = article.find("span","price-new")
        else:
            prix = article.find("span","price-regular")
        booksheet.write(row,0,product.get_text(' | ','br/'))
        booksheet.write(row,1,prix.get_text())
        row = row+1
        print(product.get_text(' | ','br/'))
        print(prix.get_text())


    for num in {"2","3","4"}:
        pages = "I-Page"+num+"_48"
        html = urlopen("https://www.nocibe.fr/yves-saint-laurent/C-47287/"+pages+"?allProduct=true")
        bsObj = BeautifulSoup(html, "html.parser")
        try:
            bsObj_list = bsObj.find("div", "container-rwd under-fixed is-banner").find_all("article",limit=50)
        except AttributeError:
            print("页面缺少一些属性！不过不用担心！")
            
        for article in bsObj_list:
            product = article.find("a",class_="brown one-product")
            if (article.find("span","price-regular") == None):
                prix = article.find("span","price-new")
            else:
                prix = article.find("span","price-regular")
            booksheet.write(row,0,product.get_text(' | ','br/'))
            booksheet.write(row,1,prix.get_text())
            row = row+1
            print(product.get_text(' | ','br/'))
            print(prix.get_text())
    workbook.save('/Users/alienware/Desktop/YSL_excel.xls')  


Yating_deng()

