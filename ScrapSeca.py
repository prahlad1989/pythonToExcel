# mugshots scraper
# runs in python 3
# tested on macbook air and macbook pro
# running macos high sierra


# python dependencies
# these dont need any installation
# os is used for doing directory operations
import os
# sys is used for retrieving arguments from the terminal
import sys
# time is used for making the script pause and wait
import time
# string is used for string operations
import string
# TODO argparse is used for retrieving arguments from the terminal
import argparse
# pip install urllib3
import urllib.request
# pip install selenium
from collections import OrderedDict

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
import re
import xlwt
from xlwt import Workbook

# pip install Pillow
import PIL.Image

# execute this file script.py with python 3
# to check the version of python you can open terminal and write
# python --version
# or
# python3 --version
# and it will tell you if you are using python 3

# in this machine my alias is python3, so the command is
# python3 script.py "subject to scrape" maxImages
# where
# script.py is this file, can be retrieved with sys.argv[0]
# "subject to scrape" is a string,  retrievable with sys.argv[1]
# maxImages is a number, retrievable with sys.argv[2]

# retrieve the script


# declaration of scraping function
def scrapeMugshots():



    # open a new google chrome window
    driver = webdriver.Chrome()

    #declare actionChains for right-click
    actionChains = ActionChains(driver)

    # set the window size
    driver.set_window_size(1200, 800)

    # queryBegin for google images search
    queryURL = "https://www.seca.ch/Membership/Members.aspx"

    # do the query
    print("query mugshots")

    # do the query
    driver.get(queryURL)

    # wait for loading
    time.sleep(1.0)

    # find div with page related classes
    pageResult = driver.find_elements_by_xpath("//nav/[@class='pages']").text

    # find div with pageresults   #Displaying results 1-25 (of 546)
    pageResult = driver.find_element_by_class_name("PagerResults").text
    pageSizeInfo = re.findall("\s\d-\d+\s", pageResult)[0].strip()
    print("page size info {0}".format(pageSizeInfo))
    pageSize = pageSizeInfo.split("-")  # 1-25  --> 25
    pageSize= int(pageSize[1])

    totalRecords = int(re.findall("of\s(\d+)", pageResult)[0].strip())
    print("totalRecords {0}".format(totalRecords))

    numOfPages = int(totalRecords/pageSize)
    if(totalRecords%pageSize !=0):
        numOfPages+=1
    time.sleep(0.2)
    queryURL = "https://www.seca.ch/Membership/Members.aspx?page={0}"
    allRows =list()
    for pageNum in range(1,1+1):
        print("page number is{0}".format(pageNum))
        driver.get(queryURL.format(pageNum))
        time.sleep(0.2)
        items = driver.find_elements_by_class_name("default_list_member_item")
        pageLinks = list(map(lambda x:x.find_element_by_tag_name("a").get_attribute("href"), items))
        for eachPageLink in pageLinks:
            driver.get(eachPageLink)
            time.sleep(0.2)
            eachRowDict =OrderedDict()
            eachRowDict['Name'] = driver.find_elements_by_xpath("//div[@class='content_middle']/h1/p")[0].text
            eachRowDict['Street'] = driver.find_elements_by_xpath("//div[@class='content_left']/div[@class='tabs']//table//td")[0].text.split("\n")[0]
            allRows.append(eachRowDict)

    wb = Workbook()

    # add_sheet is used to create sheet.
    sheet1 = wb.add_sheet('SecaScrapped')
    sheet1.write(0, 0, 'Name')
    sheet1.write(0, 1, 'Street')
    for i in range(len(allRows)):
        eachRow = list(allRows[i].values())
        sheet1.write(i + 1, 0, eachRow[0])
        sheet1.write(i + 1, 1, eachRow[1])

    wb.save('SecaScrapped.xls')



    # wait for loading


    # last sleep to make sure that we see whats going on
    time.sleep(5)
#
# def convertToJPG():
#
#     # change directory if needed
#     if os.path.exists(folderName):
#         os.chdir(folderName)
#
#     # try to create new folder to store jpg converted files
#     try:
#         os.makedirs("_jpg")
#         print("_jpg folder created")
#     except:
#         print("_jpg folder already exists, moving on")
#
#     for filename in os.listdir(os.getcwd()):
#         # open original image
#         if (filename[0] == "i"):
#             try:
#                 PIL.Image.open(filename).convert("RGB").save("_jpg/" + filename + ".jpg", quality=100)
#                 print("converted " + filename + " to jpg")
#             except:
#                 print("could not convert " + filename)

# call the function for scraping
scrapeMugshots()

# convert all the images to JPG format
#convertToJPG()

# final message, goodbye, the end
print("finished scraping yay")
