import os
from selenium import webdriver
from pyhtml2pdf import converter
PATH_DRIVER = "chromedriver.exe"

current_working_directory = os.getcwd()
chrome_options = webdriver.ChromeOptions()
prefs = {'download.default_directory': current_working_directory}
chrome_options.add_experimental_option('prefs', prefs)
driver = webdriver.Chrome(PATH_DRIVER, options=chrome_options)

path = os.path.abspath('test.html')
converter.convert(f'file:///{path}', 'sample.pdf')