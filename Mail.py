import selenium
import pandas as pd
import clipboard
import pdfcrowd
import shutil
import os
import sys
import time
import smtplib
import mimetypes
import xlrd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from zipfile import ZipFile
from pyhtml2pdf import converter
from email.message import EmailMessage

if (__name__ == "__main__"):
    PATH_DRIVER = "chromedriver.exe"
    PATH_EXCEL = "test.xlsx"

    xlrd.xlsx.ensure_elementtree_imported(False, None)
    xlrd.xlsx.Element_has_iter = True

    current_working_directory = os.getcwd()
    client = pdfcrowd.HtmlToPdfClient('TheDifferent', 'ce544b6ea52a5621fb9d55f8b542d14d')
    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory': current_working_directory}
    chrome_options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(PATH_DRIVER, options=chrome_options)

    data = pd.DataFrame(pd.read_excel(PATH_EXCEL),columns=['Alıcı', 'E-Posta'])
    data_Alici = (data['Alıcı'] + " " + data['Alıcı']).tolist()
    data_Alici_pdf = (data['Alıcı']).tolist()
    print(data_Alici_pdf)
    for i in range(len(data_Alici)):
        data_Alici[i] = data_Alici[i].split()
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
                calendar_start = WebDriverWait(driver, 20).until(
                    EC.presence_of_all_elements_located((By.XPATH, "/html/body/div[1]/div/div[2]/div/div/div/div/div/div[2]/div[2]/div/div/div/div/div[2]"
                                                               "/div/div/div[2]/div[2]/div/div/div[1]/div/div/div/div[1]/div/div/div/div[1]/div[2]/div/img"))
                )
                clipboard.copy(input("Başlangıç Tarihi Giriniz(Örn: 08/09/2020) : "))
                calendar_start[0].click()
                date_start = driver.find_element_by_id("date-gen__1024")
                date_start.send_keys(clipboard.paste())

                calender_end = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div/div/div/div/div[2]/div[2]/div/div/div/div/div[2]/"
                                                    "div/div/div[2]/div[2]/div/div/div[1]/div/div/div/div[1]/div/div/div/div[2]/div[2]/div/img")
                clipboard.copy(input("Bitiş Tarihi Giriniz (Örn: 10/09/2020) : "))
                calender_end.click()
                date_end = driver.find_element_by_id("date-gen__1025")
                date_end.send_keys(clipboard.paste())

                submit_button = driver.find_element_by_id("gen__1026").click()
                time.sleep(5)

                bill_number = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div/div/div/div/div[2]/div[2]"
                                                       "/div/div/div/div/div[2]/div/div/div[2]/div[2]/div/div/div[2]/div/div/table/tbody/tr[11]/td/span").text
                bill_number = bill_number.split()

                page_number = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div/div/div/div/div[2]/div[2]/div/div/div/div/div[2]"
                                                       "/div/div/div[2]/div[2]/div/div/div[2]/div/div/table/tbody/tr[11]/td/div/span[5]").text
                page_number = page_number.replace("/", "")

                bill_index = 0

                for page_index in range(0, int(page_number)):
                    checkbox = driver.find_elements_by_class_name("csc-table-select")

                    for checkbox_index in range(1, len(checkbox)):
                        approvebox = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div/div/div/div/div[2]/div[2]/div/div/div/div/div[2]"
                                                              "/div/div/div[2]/div[2]/div/div/div[2]/div/div/table/tbody/tr[" + str(checkbox_index) + "]/td[7]/a/i").get_attribute("class")
                        person_name = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div/div/div/div/div[2]/div[2]/div/div/div/div/div[2]/div/div/div[2]/"
                                                               "div[2]/div/div/div[2]/div/div/table/tbody/tr[" + str(checkbox_index) + "]/td[4]/span").text
                        person_name = person_name.split()

                        if (approvebox == "fa fa-check" and person_name == data_Alici[bill_index]):
                            checkbox[checkbox_index].click()

                            download = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div/div/div/div/div[2]/div[2]/div/div/div/div/div[2]/"
                                                            "div/div/div[2]/div[2]/div/div/div[1]/div/div/div/div[2]/div/div/div/div[5]/div/div/input").click()
                            checkbox[checkbox_index].click()

                            time.sleep(5)

                            for zip_item in os.listdir(current_working_directory):
                                if zip_item.endswith(".zip"):
                                    with ZipFile(zip_item, 'r') as zipObj:
                                        zipObj.extractall()
                                        zipObj.close()
                                        os.remove(zip_item)

                            for html_file in os.listdir(current_working_directory):
                                if html_file.endswith(".html"):
                                    path = os.path.abspath('test.html')
                                    converter.convert(f'file:///{path}', 'sample.pdf')
                                    os.remove(html_file)

                            for xml_file in os.listdir(current_working_directory):
                                if xml_file.endswith(".xml"):
                                    os.remove(xml_file)


                            recipient = data_Mail[bill_index]
                            message = EmailMessage()
                            sender = "emirhandemir922@gmail.com"
                            password = 'E121011e'

                            message['From'] = sender
                            message['To'] = recipient
                            message['Subject'] = 'Learning to send email from medium.com'
                            body = 'Hello I am learning to send emails using Python!!!'

                            message.set_content(body)
                            mime_type, _ = mimetypes.guess_type(data_Alici_pdf[bill_index] + '.pdf')
                            mime_type, mime_subtype = mime_type.split('/')

                            with open(data_Alici_pdf[bill_index] + '.pdf', 'rb') as file:
                                message.add_attachment(file.read(),
                                                   maintype=mime_type,
                                                   subtype=mime_subtype,
                                                   filename=data_Alici_pdf[bill_index] + '.pdf')

                            mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
                            mail_server.login(sender, password)
                            mail_server.send_message(message)

                            bill_index = bill_index + 1

                            time.sleep(1)

                    next_page = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div/div/div/div/div[2]/div[2]/div/"
                                                     "div/div/div/div[2]/div/div/div[2]/div[2]/div/div/div[2]/div/div/table/tbody/tr[11]/td/div/span[6]")
                    next_page.click()
                    time.sleep(1)

                mail_server.quit()

            finally:
                driver.refresh()
        finally:
            driver.refresh()
    finally:
        driver.refresh()

    print("İşleminiz bitmiştir.")