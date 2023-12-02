from lxml import etree
from os import path
import xlwings as xw
import requests
import re
import os
import zipfile


class Manager:
    def __init__(self, modFilePath):
        self.filename2Simple = {}
        self.filename2Real = {}

        files = os.listdir(modFilePath)
        num = 1
        xb = xw.Book()
        xs = xb.sheets["Sheet1"]
        xs.range(num, 2).value = "Mod文件名"
        xs.range(num, 3).value = "Mod信息页名"
        xs.range(num, 4).value = "Mod信息"
        xs.range(num, 5).value = "Mod搜索名"
        xs.range(num, 6).value = "是否为Forge Mod"

        for file in files:
            num += 1
            fileName = path.basename(file)
            isForgeMod = self.isForgeMod(fileName)
            modName = self.getModName(fileName)
            if not modName:
                modName = self.simplifyName(fileName)
            self.filename2Simple[fileName] = modName
            self.loadSearchWeb(modName)
            searchResultDic = self.getModname2UrlDic()

            if not searchResultDic:
                xs.range(num, 1).value = num - 1
                xs.range(num, 2).value = fileName
                xs.range(num, 3).value = "查无此mod"
                xs.range(num, 4).value = "\\"
                xs.range(num, 5).value = modName
                xs.range(num, 6).value = isForgeMod
                print(fileName, "->", modName, ": 查无此mod")
            else:
                modInfoName = list(searchResultDic.keys())[0]
                modInfoSide = self.isServerNeeded(list(searchResultDic.values())[0])

                xs.range(num, 1).value = num - 1
                xs.range(num, 2).value = fileName
                xs.range(num, 3).value = modInfoName
                xs.range(num, 4).value = modInfoSide
                xs.range(num, 5).value = modName
                xs.range(num, 6).value = isForgeMod
                print(fileName, "->", modName, "->", modInfoName, ":", modInfoSide)
        xb.save("result.xlsx")

    def getModName(self, modFileName):
        if re.search(r"\.jar$", modFileName):  # Check if is jar file
            with zipfile.ZipFile("./mods/" + modFileName, 'r') as jarfile:
                infiles = jarfile.namelist()
                for infile in infiles:
                    if re.search(r"mods.toml$", infile):
                        content = jarfile.read(infile).decode("UTF-8", errors="ignore")
                        modName = re.search(r"displayName=\"(.+)\"", content)
                        if modName:
                            return modName.group(1)
        return None

    def simplifyName(self, name):
        finder = re.compile(r"^([a-zA-Z'_]*)?(【(.*)】)?([a-zA-Z'_]*)")
        findResult = finder.search(name)
        if findResult.group(3) is None:
            return findResult.group(1)
        else:
            if findResult.group(3) == "前置":
                return findResult.group(4)
            else:
                return findResult.group(4)

    def loadSearchWeb(self, modName):
        params = {
            "key": modName,
            "filter": 1,
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36 Edg/109.0.1518.70"
        }
        response = requests.get("https://search.mcmod.cn/s", params=params)
        with open("searchWeb.html", "w", encoding="UTF-8") as f:
            f.write(response.text)

    def getModname2UrlDic(self, searchWebSrc="searchWeb.html"):
        parser = etree.HTMLParser(recover=True, encoding="UTF-8")
        tree = etree.parse(searchWebSrc, parser=parser)
        elements = tree.xpath("//div[@class='result-item']/div[@class='head']/a[@target='_blank']")
        # for element in elements:
        #     # d = etree.tostring(element, encoding="UTF-8").decode("UTF-8")
        #
        #     # print(element.get("href"))
        #     # print('*' * 50)
        # for element in elements:
        # print(' '.join(etree.tostring(element, method="text", encoding="UTF-8").decode("UTF-8").split()))
        # print('*' * 50)
        # print(etree.tostring(tree, encoding="UTF-8").decode("UTF-8"))
        modname2UrlDic = {}
        for element in elements:
            modname2UrlDic[
                ' '.join(
                    etree.tostring(element, method="text", encoding="UTF-8").decode("UTF-8").split())] = element.get(
                "href")
        # print(modname2UrlDic)
        # print('-' * 50)
        return modname2UrlDic

    def isServerNeeded(self, modUrl):
        params = {
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36 Edg/109.0.1518.70"
        }
        response = requests.get(modUrl)
        # Store the mod web file
        with open("modWeb.html", "w", encoding="UTF-8") as f:
            f.write(response.text)
        # Parse the mod web file
        parser = etree.HTMLParser(recover=True, encoding="UTF-8")
        tree = etree.parse("modWeb.html", parser=parser)
        elements = tree.xpath("//div[@class='class-info']//ul[@class='col-lg-12']/li[@class='col-lg-4']")

        for element in elements:
            str = etree.tostring(element, encoding="UTF-8", method="text").decode("UTF-8")
            if re.search("服务端需装", str):
                return "服务端需装"
            if re.search("服务端无效", str):
                return "服务端无效"
            if re.search("服务端可选", str):
                return "服务端可选"
        return "未知" + '[' + modUrl + ']'

    def isForgeMod(self, fileName):
        with zipfile.ZipFile("./mods/" + fileName, 'r') as jarfile:
            infiles = jarfile.namelist()
            for infile in infiles:
                if re.search(r"mods.toml$", infile):
                    return True
        return False

if __name__ == '__main__':
    Manager("./mods")
