import os
from os import path
from bs4 import BeautifulSoup
import urllib.request as urllib
import pandas as pd
import xlwt
import xlrd
import datetime
import requests
from time import sleep

def getASINfromMnrate(url):
    headers = {'User-agent': 'Mozilla/5.0'}
    result = []
    # html = urllib.urlopen(url)
    r = requests.get(url,headers=headers)
    html = r.content
    soup = BeautifulSoup(html, "lxml")
    for tag in soup.find_all('a', attrs={"class": "original_link"}):
        href = tag.get("href")
        if href.find("https://mnrate.com/item/aid/") > -1:
            href = href.split("/")
            href = href[-1]
            print(href)
            if len(href) is 10:
                if href not in result:
                    result.append(href)
    return result


def getASINfromAmazon(url):
    headers = {'User-agent': 'Mozilla/5.0'}
    result = []
    # html = urllib.urlopen(url)
    opener = urllib.build_opener()
    opener.addheaders = [('User-agent', 'Mozilla/5.0')]
    response = opener.open(url)
    html = response.read() 
    # r = requests.get(url)
    # html = r.content
    soup = BeautifulSoup(html, "lxml")
    # print(soup)
    for tag in soup.find_all('a', attrs={"class": "a-link-normal"}):
        # print(tag)
        href = tag.get("href")
        print(href)
        if href.find("/dp/") > -1:
            href = href.split("/")
            if len(href) >= 6:
                asinCode = href[5]
                print(asinCode)
                if len(asinCode) is 10:
                    if asinCode not in result:
                        print(asinCode)
                        result.append(asinCode)
    return result


def main():
    inputpath = "./inputUrl.xlsx"
    inputpath2 = "./inputUrl.xls"
    inputUrl = pd.DataFrame()
    if os.path.isfile(inputpath):
        inputUrl = pd.read_excel(inputpath)
    elif os.path.isfile(inputpath2):
        inputUrl = pd.read_excel(inputpath)
    else:
        print("inputUrl.xlsx (またはinputUrl.xls) ファイルが存在しません！")
        return
    columns = inputUrl.columns
    if 'amazon' not in columns:
        print('inputUrlにはファイルにはamazonコラムが入っていません。')
        return
    if 'mnrate' not in columns:
        print('inputUrlにはファイルにはmnrateコラムが入っていません。')
        return

    amazonUrl = []
    mnrateUrl = []
    for i in range(1):
        amazonBaseUrl = "https://www.amazon.co.jp/s?i=kitchen&bbn=3828871&rh=n%3A3828871%2Cp_n_availability%3A2227307051%2Cp_36%3A150000-200000&page={}&qid=1558190853&ref=sr_pg_2".format(i+1)
        amazonUrl.append(amazonBaseUrl)
    for i in range(400):
        mnrateBaseUrl = "https://mnrate.com/search?i=Kitchen&kwd=&ex_asa%5B%5D=e&ex_asa%5B%5D=e&ex_asa%5B%5D=p&ex_asa%5B%5D=e&ex_asa%5B%5D=p&ex_asa%5B%5D=e&ex_asa%5B%5D=p&nppp_min=120&s=r&tn_min=2&tn_max=&p={}".format(i+1)
        mnrateUrl.append(mnrateBaseUrl)
    # write path
    outputPath = "./output.xls"
    amazonASIN = []
    mnrateASIN = []
    count = 0
    for link in amazonUrl:
        count = count + 1
        start = datetime.datetime.now()
        asin = getASINfromAmazon(link)
        end = datetime.datetime.now()
        print("count:{},start: {},end: {},excute: {}".format(count,start,end,end-start))
        for element in asin:
            amazonASIN.append(element)

    count = 0
    sleepTime = 300
    # for link in mnrateUrl:
    #     count = count + 1
    #     if count % 40 is 0:
    #         print("wait for {}s".format(sleepTime))
    #         sleep(sleepTime)
    #     start = datetime.datetime.now()
    #     asin = getASINfromMnrate(link)
    #     end = datetime.datetime.now()
    #     print("count:{},start: {},end: {},excute: {}".format(count,start,end,end-start))
    #     for element in asin:
    #         mnrateASIN.append(element)

    # check if output file is exist
    if os.path.isfile(outputPath):
        os.remove(outputPath)
        print("Removed {}".format(outputPath))
    # write to output.xls
    col = 0
    row = 0
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('AMAZON')
    for index in range(len(amazonASIN)):
        value = amazonASIN[index]
        # print("index: {},value: {}".format(index, value))
        if row is 1000:
            col = col + 1
            row = 0
        sheet.write(row, col, value)
        row = row + 1
    sheet2 = workbook.add_sheet('MNRATE')
    row = 0
    col = 0
    for index in range(len(mnrateASIN)):
        value = mnrateASIN[index]
        if row is 1000:
            col = col + 1
            row = 0
        sheet2.write(row, col, value)
        row = row + 1
    workbook.save(outputPath)
    print("Wrote to {}".format(outputPath))


if __name__ == "__main__":
    main()


