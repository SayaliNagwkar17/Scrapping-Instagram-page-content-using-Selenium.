from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
from selenium.common.exceptions import NoSuchElementException
import xlsxwriter

driver = webdriver.Chrome("==========")#Enter chrome driver path
driver.maximize_window()
driver.get("https://www.instagram.com/accounts/login/")

time.sleep(2)

#Enter username and password

Username = driver.find_element_by_name("username")
Username.send_keys("========")#Enter a username

Password = driver.find_element_by_name("password")
Password.send_keys("========")#SEnter a pwd

act=ActionChains(driver)
act.send_keys(Keys.ENTER).perform()
time.sleep(3)

#Selecting 'Not Now' on Save Login info page
driver.find_element_by_xpath('//*[@id="react-root"]/section/main/div/div/div/div/button').click()
time.sleep(2)

#Selecting 'Not Now' on Allow notifications page
driver.find_element_by_xpath('/html/body/div[4]/div/div/div/div[3]/button[2]').click()
time.sleep(2)

#Search keyword in search tab and open the corresponding page.

keyword='================'
driver.find_element_by_xpath("//*[@id='react-root']/section/nav/div[2]/div/div/div[2]/input").send_keys(keyword)

act=ActionChains(driver)
act.send_keys(Keys.ENTER).perform()
time.sleep(3)

driver.find_element_by_xpath('//a[contains (@href,"/'+keyword+'/")]').click()
time.sleep(2)

#Scrolling complete page.
posts = []
lenOfPage=driver.execute_script("window.scrollTo(0,document.body.scrollHeight);var lenOfPage=document.body.scrollHeight;return lenOfPage;")
match=False
while(match==False):
    lastCount=lenOfPage
    time.sleep(3)

    links = driver.find_elements_by_tag_name('a')
    for link in links:
        post = link.get_attribute('href')
        if '/p/' in post:
            posts.append(post)

    lenOfPage = driver.execute_script("window.scrollTo(0,document.body.scrollHeight);var lenOfPage=document.body.scrollHeight;return lenOfPage;")
    if lastCount==lenOfPage:
        match=True


set_post=set(posts)
new_post=list(set_post)
print("Total number of post:",len(new_post))

workbook = xlsxwriter.Workbook(keyword + ".xlsx")
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': True})

worksheet.write('A1', 'URL',bold)
worksheet.write('B1', 'Caption',bold)
worksheet.write('C1', 'Like',bold)
worksheet.write('D1', 'Comment',bold)

row = 1


for i in new_post:
    driver.get(i)
    time.sleep(2)

    worksheet.write_string(row, 0, str(i))
    #===========Caption=============#

    caption=driver.find_element_by_xpath('//*[@id="react-root"]/section/main/div/div[1]/article/div/div[3]/div[1]/ul/div/li/div/div/div[2]/span').text
    worksheet.write_string(row, 1, str(caption))
    #============Likes===============#
    try:
        likepeople = driver.find_element_by_xpath(
            '//*[@id="react-root"]/section/main/div/div[1]/article/div/div[3]/section[2]/div/div/button').text
        worksheet.write_string(row, 2,str(likepeople))
    except NoSuchElementException:
        worksheet.write_string(row, 2, str("No Likes"))

        # =============Comments============#

    try:
        while driver.find_element_by_xpath(
                "//*[@id='react-root']/section/main/div/div/article/div/div[3]/div[1]/ul/li/div/button").is_displayed():
            driver.find_element_by_xpath(
                "//*[@id='react-root']/section/main/div/div/article/div/div[3]/div[1]/ul/li/div/button").click()
            no_comment = driver.find_elements_by_tag_name('h3')

            com = []
            time.sleep(2)
            for j in range(1, len(no_comment) + 1):
                comment = driver.find_element_by_xpath(
                    "//*[@id='react-root']/section/main/div/div[1]/article/div/div[3]/div[1]/ul/ul[{}]/div/li/div/div[1]/div[2]/span".format(
                        j)).text
                com.append(comment)
            worksheet.write_string(row, 3, str(com))

    except NoSuchElementException:
        no_comment = driver.find_elements_by_tag_name('h3')

        if len(no_comment) == 0:
            worksheet.write_string(row, 3, str("No comment"))

        elif len(no_comment) == 1:
            comment = driver.find_element_by_xpath(
                "//*[@id='react-root']/section/main/div/div[1]/article/div/div[3]/div[1]/ul/ul/div/li/div/div[1]/div[2]/span").text

            worksheet.write_string(row, 3, str(comment))

        else:
            no_comment = driver.find_elements_by_tag_name('h3')
            com = []
            for j in range(1, len(no_comment) + 1):
                comment = driver.find_element_by_xpath(
                    "//*[@id='react-root']/section/main/div/div[1]/article/div/div[3]/div[1]/ul/ul[{}]/div/li/div/div[1]/div[2]/span".format(
                        j)).text
                com.append(comment)
            worksheet.write_string(row, 3, str(com))

    row += 1

workbook.close()
print("Excel file ready")

