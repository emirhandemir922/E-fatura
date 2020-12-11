import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import pandas as pd
import pdfcrowd
import os
import sys
import time

client = pdfcrowd.HtmlToPdfClient('demo', 'ce544b6ea52a5621fb9d55f8b542d14d')
PATH_DRIVER = "chromedriver.exe"
PATH_EXCEL = "Siparişleriniz.xlsx"
FILES_PATH = "%USERPROFILE%/Downloads/*.zip"
driver = webdriver.Chrome(PATH_DRIVER)


data = pd.DataFrame(pd.read_excel(PATH_EXCEL),columns=['Alıcı', 'E-Posta'])
data_Alıcı = data['Alıcı'].tolist()
data_Mail = data['E-Posta'].tolist()

driver.get("https://earsivportal.efatura.gov.tr/intragiris.html")
userid = driver.find_element_by_id("userid")
userid.send_keys("12305487")
userid.send_keys(Keys.RETURN)

userpassword = driver.find_element_by_id("password")
userpassword.send_keys("662252")
userpassword.send_keys(Keys.RETURN)

enter = driver.find_element_by_name("action")
enter.send_keys(Keys.RETURN)

try:
    dropdown = WebDriverWait(driver, 20).until(
        EC.presence_of_all_elements_located((By.ID, "gen__1006"))
    )
    dropdown = driver.find_element_by_id("gen__1006")
    dropdown = Select(dropdown)
    dropdown.select_by_visible_text("e-Arşiv Portal")

    try:
        document = WebDriverWait(driver, 20).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "cstree-closed"))
        )
        document = driver.find_element_by_class_name("cstree-closed").click()

        drafts = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div/div/div/div/"
                                              "div[2]/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/ul/li[2]/ul/li[2]/a").click()

        try:
            date_start = WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located((By.ID, "date-gen__1024"))
            )
            date_start[0].send_keys(input("Başlangıç Tarihi Giriniz(Örn: 08/09/2020) : "))

            date_end = driver.find_element_by_id("date-gen__1025")
            date_end.send_keys(input("Bitiş Tarihi Giriniz (Örn: 10/09/2020) : "))

            submit_button = driver.find_element_by_id("gen__1026").click()

            time.sleep(5)

            for bill_index in range(len(data_Alıcı)):
                checkbox = driver.find_elements_by_class_name("csc-table-select")

                for checkbox_index in range(1, len(checkbox)):
                    checkbox[checkbox_index].click()

                    download = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div/div/div/div/div[2]/div[2]/div/div/div/div/div[2]/"
                                                            "div/div/div[2]/div[2]/div/div/div[1]/div/div/div/div[2]/div/div/div/div[5]/div/div/input").click()

                    client.convertFileToFile('C:/Users/Emirhan/AppData/Local/Temp/Temp1_test.zip/3716810c-ad85-496e-b0f6-407922021ca0_f.html', data_Alıcı[bill_index] + '.pdf')



                next_page = driver.find_elements_by_class_name("csc-table-paging-btn")
                next_page[2].click()

        finally:
            driver.refresh()
    finally:
        driver.refresh()
finally:
    driver.refresh()
