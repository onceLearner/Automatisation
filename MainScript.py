import openpyxl
from MainFunctions import *

from datetime import time

from selenium import webdriver
import time
import pyautogui

from selenium.common.exceptions import TimeoutException
from selenium.webdriver import ActionChains

from selenium.webdriver.chrome.options import Options
import openpyxl



from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait



def Guru(path,start,stop,isBleed,isAdult,isColor,isSize):
    page1 = "https://kdp.amazon.com/en_US/title-setup/paperback/new/details?ref_=kdp_BS_D_cr_ti"

    options = webdriver.ChromeOptions()
    options.add_argument("user-data-dir=C:\\Users\\OnceLearner\\ChromeProfiles\\Profile 1")
    driver = webdriver.Chrome(options=options)
    # driver.get("https://kdp.amazon.com/en_US/title-setup/paperback/new/details?ref_=kdp_BS_D_cr_ti")
    driver.get(page1)
    wait = WebDriverWait(driver, 555)

    workbook=openpyxl.load_workbook(path,data_only=True)
    sheet=workbook.active
    a=start
    while start<=stop :

        time.sleep(2)
        firstPage(start,sheet,driver,isBleed,isAdult,isColor,isSize)
        start+=1
        time.sleep(1)
        driver.execute_script("window.open('{}')".format(page1))
        nb_tabs=driver.window_handles
        driver.switch_to.window(driver.window_handles[len(nb_tabs)-1])





    # driver.switch_to.window[i - 2]













