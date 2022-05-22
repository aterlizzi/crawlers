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
driver.get("https://www.spotsylvania.k12.va.us/domain/831")

# print the title
title = "StaffordSchools"

# add a sheet
sheet1 = wb.add_sheet(title)

sheet1.write(0, 0, 'Name')
sheet1.write(0, 1, "Email")
sheet1.write(0, 2, "Job")



# loop to loop through all of the pages, change depending on website.
idx = 0
idxMax = 8

for i in range(2, idxMax):
    # locate the containing elements for the classname staff.
    parentContainers = driver.find_elements(by=By.CLASS_NAME, value="staff" )
    for container in parentContainers:
        # find the name container
        nameContainer = container.find_element(by=By.CLASS_NAME, value="staffname")
        name = nameContainer.text # extract the text

        # find staff job container
        staffContainer = container.find_element(by=By.CLASS_NAME, value="staffjob")
        job = staffContainer.text


        # not all staff have email so handle errors.
        try:
            # find staff email
            emailParentContainer = container.find_element(by=By.CLASS_NAME, value="staffemail")
            emailLinkUnparsed = emailParentContainer.find_element(by=By.TAG_NAME, value="a").get_attribute('href') # returns string in form of "mailto: ..."
            email = emailLinkUnparsed.split("mailto:")[1] # removes the "mailto: " and only keeps the email
        except:
            email = ""

        sheet1.write(idx+1, 0, name)
        sheet1.write(idx+1, 1, email)
        sheet1.write(idx+1, 2, job)

        # increment for the xl sheet
        idx+=1

    # find the next page button, click, and sleep.
    pageButtons = driver.find_elements(by=By.CLASS_NAME, value="ui-page-number")
    # loop through buttons
    for pageButton in pageButtons:
        button = pageButton.find_element(by=By.TAG_NAME, value="a")
        pageNum = button.text
        # if button number is equal to the next button, click that button
        if(i != idxMax and pageNum == str(i+1)):
            print('found')
            pageButton.click()
            break

    # allow time for page to load
    time.sleep(5)

    # save the xl
    wb.save("ColonialForgeHighSchoolContacts.xls")

# close the webpage
driver.close()

