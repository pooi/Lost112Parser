# -*- coding:utf-8 -*-
# pip3 install xlsxwriter
# pip3 install selenium

try:
    import urllib.request as urllib2
except ImportError:
    import urllib2
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
# from fake_useragent import UserAgent
import xlsxwriter
import time
import os, sys

def checkAlert(driver):
    try:
        WebDriverWait(driver, 0).until(EC.alert_is_present(),
                                        'Timed out waiting for PA creation ' +
                                        'confirmation popup to appear.')

        alert = driver.switch_to.alert
        alert.accept()
        # print("alert accepted")
    except TimeoutException:
        pass
        # print("no alert")

def getDetailInfo(driver, js):
    driver.execute_script(js)
    time.sleep(2)
    checkAlert(driver)

    html = driver.page_source
    soup = BeautifulSoup(html)
    finds = soup.find("div", {"class": "find_info"})
    details = finds.find_all("li")

    for detail in details:
        title = detail.find('p', {"class": "find01"}).text
        if title == '물품분류':
            category = detail.find('p', {"class": "find02"}).text
            driver.back()
            time.sleep(2)
            return category

start = 1
finish = 10

start = int(input("Start Page : "))
finish = int(input("Finish Page : "))

# ua = UserAgent()
dir = os.getcwd() + '/chromedriver'
print(dir)
driver = webdriver.Chrome(dir)

# driver = webdriver.Chrome()
driver.get('https://www.lost112.go.kr/find/findList.do')

# jsList = []
# categories = []

workbook = xlsxwriter.Workbook('lost(%d-%d).xlsx' % (start, finish))
worksheet = workbook.add_worksheet()

row = 0
col = 0

try:

    dataList = []
    for page in range(start, finish+1):

        print("Page %d/%d" % (page, finish))

        driver.execute_script('fn_find_link_page(%d);' % page)
        time.sleep(2)
        checkAlert(driver)

        html = driver.page_source
        soup = BeautifulSoup(html)
        lostList = soup.find_all("td", {"class": "board_title1 "})


        for lost in lostList:
            # print(lost)
            title = lost.find('a').text
            title = title.replace('\n', '').replace('\t', '')
            js = lost.find('a')['href']
            category = getDetailInfo(driver, js)
            js = str(js)
            itemId = js[js.index("(\'")+2:js.index("\',")]
            # dataList.append(data)
            category = str(category)
            categories = category.split('>')
            c1 = categories[0].strip()
            c2 = ""
            if len(categories) > 1:
                c2 = categories[1].strip()
            data = {'id': itemId, 'title': title, 'js': js, 'category1': c1, 'category2': c2}
            print(title)

            worksheet.write(row, 0, data['id'])
            worksheet.write(row, 1, data['title'])
            worksheet.write(row, 2, data['category1'])
            worksheet.write(row, 3, data['category2'])
            worksheet.write(row, 4, data['js'])
            row += 1

finally:
    workbook.close()
    driver.close()

# print()
# for data in dataList:
#     print(data)
#     worksheet.write(row, 0, data['id'])
#     worksheet.write(row, 1, data['title'])
#     worksheet.write(row, 2, data['category'])
#     worksheet.write(row, 3, data['js'])
#     row += 1

# workbook.close()

# driver.execute_script(jsList[0])
# time.sleep(3)
# print(getDetailInfo(driver))
# html = driver.page_source
# soup = BeautifulSoup(html)
# lostList = soup.find_all("p", {"class": "find_info_name"})
#
# print(lostList)

# driver.close()