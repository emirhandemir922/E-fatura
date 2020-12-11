import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import pandas as pd
from num2words import num2words
import time

PATH_DRIVER = "chromedriver.exe"
PATH_EXCEL = "Siparişleriniz.xlsx"
driver = webdriver.Chrome(PATH_DRIVER)

data = pd.DataFrame(pd.read_excel(PATH_EXCEL),columns=['Alıcı', 'Fatura Adresi', 'Adet', 'Birim Fiyatı', 'Faturalanacak Tutar'])
data_Alıcı = data['Alıcı'].tolist()
data_Adres = data['Fatura Adresi'].tolist()
data_Adet = data['Adet'].tolist()
data_BirimFiyatı = (data['Birim Fiyatı'] / 1.08).tolist()

driver.get("https://earsivportal.efatura.gov.tr/intragiris.html")
userid = driver.find_element_by_id("userid")
userid.send_keys("12305487")
userid.send_keys(Keys.RETURN)

userpassword = driver.find_element_by_id("password")
userpassword.send_keys("662252")
userpassword.send_keys(Keys.RETURN)

enter = driver.find_element_by_name("action")
enter.send_keys(Keys.RETURN)

for index in range(len(data_Alıcı)):
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

            bill = driver.find_element_by_link_text("5000/30.000TL Fatura Oluştur").click()

            try:
                TCKN = WebDriverWait(driver, 20).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//*[@id='gen__1033']"))
                )

                TCKN[0].send_keys("11111111111")
                TCKN[0].send_keys(Keys.RETURN)

                time.sleep(1)

                name = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div/div/div/div/div[2]/div[2]"
                                                       "/div/div/div/div/div[2]/div/div/div[2]/div[2]/div/div/div[1]/div/div/div/div[2]/div/div/fieldset/table/tr[3]/td[2]/input")
                name.send_keys(data_Alıcı[index])
                name.send_keys(Keys.RETURN)

                time.sleep(1)

                lastname = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div/div/div/div/div[2]/div[2]"
                                                           "/div/div/div/div/div[2]/div/div/div[2]/div[2]/div/div/div[1]/div/div/div/div[2]/div/div/fieldset/table/tr[4]/td[2]/input")
                lastname.send_keys(data_Alıcı[index])
                lastname.send_keys(Keys.RETURN)

                time.sleep(1)

                country = driver.find_element_by_id("gen__1042-i")
                country.send_keys("Türkiye")
                country.send_keys(Keys.RETURN)

                adress = driver.find_element_by_class_name("csc-textarea")
                adress.send_keys(data_Adres[index])
                adress.send_keys(Keys.RETURN)

                createbill = driver.find_element_by_id("gen__1093").click()

                product = driver.find_element_by_id("gen__1148")
                product.send_keys("Pijama")
                product.send_keys(Keys.RETURN)

                quantity = driver.find_element_by_id("gen__1149")
                quantity.send_keys(data_Adet[index])
                quantity.send_keys(Keys.RETURN)

                unit = driver.find_element_by_id("gen__1150")
                unit = Select(unit)
                unit.select_by_visible_text("Adet")

                data_BirimFiyatı[index] = str(data_BirimFiyatı[index]).replace(".", ",")
                unitcost = driver.find_element_by_id("gen__1151")
                unitcost.clear()
                unitcost.send_keys(data_BirimFiyatı[index])
                unitcost.send_keys(Keys.RETURN)

                KDV = driver.find_element_by_id("gen__1158")
                KDV = Select(KDV)
                KDV.select_by_visible_text("8")

                data_FaturalanacakTutar = num2words(data['Faturalanacak Tutar'][index], to='currency', lang='tr')
                exp = driver.find_element_by_id("gen__1109")
                exp.send_keys(data_FaturalanacakTutar.upper())
                exp.send_keys(Keys.RETURN)

                createbutton = driver.find_element_by_id("gen__1112").click()

                time.sleep(1)

                okbutton = driver.find_element_by_xpath("/html/body/div[6]/div[2]/div/div/div/input").click()

            finally:
                driver.refresh()
        finally:
            driver.refresh()
    finally:
        driver.refresh()
