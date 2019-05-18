import os
from os import path
from bs4 import BeautifulSoup
import urllib.request as urllib
import pandas as pd
import xlwt
import xlrd


def getASINfromMnrate(url):
    result = []
    html = urllib.urlopen(url)
    soup = BeautifulSoup(html, "lxml")
    for tag in soup.find_all('a', attrs={"class": "original_link"}):
        href = tag.get("href")
        if href.find("https://mnrate.com/item/aid/") > -1:
            href = href.split("/")
            href = href[-1]
           if len(href) is 10:
                    if href not in result:
                        result.append(href)
    return result


def getASINfromAmazon(url):
    result = []
    html = urllib.urlopen(url)
    soup = BeautifulSoup(html, "lxml")
    for tag in soup.find_all('a', attrs={"class": "a-link-normal"}):
        # print(tag)
        href = tag.get("href")
        if href.find("/dp/") > -1:
            href = href.split("/")
            if len(href) >= 6:
                asinCode = href[5]
                if len(asinCode) is 10:
                    if asinCode not in result:
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

    amazonUrl = inputUrl['amazon']
    mnrateUrl = inputUrl['mnrate']

    # write path
    outputPath = "./output.xls"
    amazonASIN = []
    mnrateASIN = []
    count = 0
    for link in amazonUrl:
        if count is 400:
            break
        count = count + 1
        asin = getASINfromAmazon(link)
        for element in asin:
            amazonASIN.append(element)

    count = 0
    for link in mnrateUrl:
        if count is 1000:
            break
        count = count + 1
        asin = getASINfromMnrate(link)
        for element in asin:
            mnrateASIN.append(element)

    # check if output file is exist
    if os.path.isfile(outputPath):
        os.remove(outputPath)
        print("Removed {}".format(outputPath))
    # write to output.xls
    col = 0
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('AMAZON')
    for index in range(len(amazonASIN)):
        value = amazonASIN[index]
        # print("index: {},value: {}".format(index, value))
        if index > 0:
            if index % 1000 is 0:
                col = col + 1
        sheet.write(index, col, value)
    sheet2 = workbook.add_sheet('MNRATE')
    for index in range(len(mnrateASIN)):
        value = mnrateASIN[index]
        if index > 0:
            if index % 1000 is 0:
                col = col + 1
        sheet2.write(index, col, value)
    workbook.save(outputPath)
    print("Wrote to {}".format(outputPath))


if __name__ == "__main__":
    main()


