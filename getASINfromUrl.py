import os
from os import path
from bs4 import BeautifulSoup
import pandas as pd
import xlwt
import xlrd
import datetime
import requests
from time import sleep

def getASINfromMnrate(url):
    headers = {'User-agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/36.0.1985.143 Safari/537.36'}
    result = []
    r = requests.get(url,headers=headers)
    html = r.content
    # print(len(html))
    while(len(html)<400000):
        sleep(180)
        r = requests.get(url,headers=headers)
        html = r.content
        # print("request again for url {}".format(url))
    soup = BeautifulSoup(html, "lxml")
    countASIN = 0
    for tag in soup.find_all('a', attrs={"class": "original_link"}):
        href = tag.get("href")
        if href.find("https://mnrate.com/item/aid/") > -1:
            href = href.split("/")
            href = href[-1]
            if len(href) is 10:
                if href not in result:
                    countASIN += 1
                    result.append(href)
    print("{} ASIN コード取得しました。".format(countASIN))
    return result


def getASINfromAmazon(url):
    headers = {'User-agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/36.0.1985.143 Safari/537.36'}
    result = []
    # try:
    #     html = urllib.urlopen(url)
    # except Exception:
    #     print("error at {}".format(url))
    # opener = urllib.build_opener()
    # opener.addheaders = [('User-agent', 'Mozilla/5.0')]
    # response = opener.open(url)
    # html = response.read() 
    r = requests.get(url,headers=headers)
    html = r.content
    while(len(html)<30000):
        # sleep(20)
        r = requests.get(url,headers=headers)
        html = r.content
        # print("request again for url {}".format(url))
    soup = BeautifulSoup(html, "lxml")
    exist = soup.find_all('div',attrs={"class": "sg-col-20-of-24 s-result-item sg-col-0-of-12 sg-col-28-of-32 sg-col-16-of-20 sg-col sg-col-32-of-36 sg-col-12-of-16 sg-col-24-of-28"})
    countAsin = 0
    for item in exist:
        dataAsin = item.attrs["data-asin"]
        if len(dataAsin) is 10:
            if dataAsin not in result:
                countAsin = countAsin + 1
                result.append(dataAsin)
    print("{} ASIN コード取得しました。".format(countAsin))
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
    amazonBaseFromExcel = inputUrl['amazon'].iloc[0]
    mnrateBaseFromExcel = inputUrl['mnrate'].iloc[0]
    amazonUrl = []
    mnrateUrl = []
    for i in range(400):
        amazonBaseUrl = amazonBaseFromExcel.format(i+1)
        amazonUrl.append(amazonBaseUrl)
    for i in range(1000):
        mnrateBaseUrl = mnrateBaseFromExcel.format(i+1)
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
        print("番号:{},開始: {},終了: {},実行時間: {}秒".format(count,start,end,end-start))
        for element in asin:
            amazonASIN.append(element)

    count = 0
    sleepTime = 300
    for link in mnrateUrl:
        count = count + 1
        start = datetime.datetime.now()
        asin = getASINfromMnrate(link)
        end = datetime.datetime.now()
        print("番号:{},開始: {},終了: {},実行時間: {}秒".format(count,start,end,end-start))
        for element in asin:
            mnrateASIN.append(element)

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
        if row == 1000:
            col = col + 1
            row = 0
        sheet.write(row, col, value)
        row = row + 1
    sheet2 = workbook.add_sheet('MNRATE')
    row = 0
    col = 0
    for index in range(len(mnrateASIN)):
        value = mnrateASIN[index]
        if row == 1000:
            col = col + 1
            row = 0
        sheet2.write(row, col, value)
        row = row + 1
    workbook.save(outputPath)
    print("Wrote to {}".format(outputPath))


if __name__ == "__main__":
    main()


