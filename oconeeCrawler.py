from selenium import webdriver
from selenium.webdriver.common.by import By
from xlwt import Workbook
import time

# create a new xl
wb = Workbook()

# initilization
PATH = "/Users/aidanterlizzi/Desktop/Main/Coding/Web Crawler/chromedriver"
driver = webdriver.Chrome(PATH)

# open the webpage
driver.get("https://www.oconeeschools.org/staff")

# print the title
title = "OconeeCounty"

# add a sheet
sheet1 = wb.add_sheet(title)

sheet1.write(0, 0, 'Name')
sheet1.write(0, 1, "Job")
sheet1.write(0, 2, "Email")



# loop to loop through all of the pages, change depending on website.
idx = 0
idxMax = 300

button = driver.find_element(by=By.ID, value="minibaseSubmit177").click()
time.sleep(3)

for i in range(1, idxMax):

    evenRows = driver.find_elements(by=By.CLASS_NAME, value="sw-flex-item-group")
    for row in evenRows:
        idx+=1
        datas = row.find_elements(by=By.TAG_NAME, value="td")
        for rowIdx, data in enumerate(datas):                
            if rowIdx == 1:
                textList = data.text.split("\n")
                sheet1.write(idx, 0, textList[0])
                sheet1.write(idx, 1, textList[1])
            if rowIdx == 3:
                link = data.find_element(by=By.TAG_NAME, value="a")
                sheet1.write(idx, 2, link.text)
        

    # find the next page button, click, and sleep.
    pageButtons = driver.find_elements(by=By.CLASS_NAME, value="ui-page-number")

    # loop through buttons
    for pageButton in pageButtons:
        button = pageButton.find_element(by=By.TAG_NAME, value="a")
        pageNum = button.text

        # find the classname to check if it is return to group.
        className = pageButton.get_attribute('class')
        classArr = className.split(" ")
        if 'ui-prev-group' in classArr:
            continue

        # if button number is equal to the next button, click that button
        if(pageNum == str(i+1) or pageNum == '...'):
            print('found')
            pageButton.click()
            break
    
    # allow time for page to load
    time.sleep(2)

    # save the xl
    wb.save("OconeeCounty.xls")

# close the webpage
driver.close()

