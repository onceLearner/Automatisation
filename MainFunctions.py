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

page1="https://kdp.amazon.com/en_US/title-setup/paperback/new/details?ref_=kdp_BS_D_cr_ti"



def firstPage(i,sheet,driver,isBleed,isAdult,isColor,isSize):
    # options = webdriver.ChromeOptions()
    # options.add_argument("user-data-dir=C:\\Users\\OnceLearner\\ChromeProfiles\\Profile 1")
    # driver = webdriver.Chrome(options=options)
    # # driver.get("https://kdp.amazon.com/en_US/title-setup/paperback/new/details?ref_=kdp_BS_D_cr_ti")
    # driver.get(page1)
    wait = WebDriverWait(driver, 755)
    ############ Page 1   ####################
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="data-print-book-title"]').send_keys(sheet["A"+str(i)].value)
    driver.find_element_by_xpath('//*[@id="data-print-book-subtitle"]').send_keys(sheet["B"+str(i)].value)
    driver.find_element_by_xpath('//*[@id="data-print-book-primary-author-first-name"]').send_keys(sheet["C" + str(i)].value)
    driver.find_element_by_xpath('//*[@id="data-print-book-primary-author-last-name"]').send_keys(sheet["D" + str(i)].value)
    driver.find_element_by_xpath('//*[@id="data-print-book-description"]').send_keys(sheet["E"+str(i)].value)


    # if public domaine or not
    public_domain="" if sheet["F"+str(i)].value else "non-"
    driver.find_element_by_xpath('//*[@id="{}public-domain"]'.format(public_domain)).click()

    # #keywords
    J=0
    for letter in "GHIJKLM":
        driver.find_element_by_xpath('//*[@id="data-print-book-keywords-' + str(J) + '"]').send_keys(sheet[letter+str(i)].value)
        J+=1
    # open categories
    driver.find_element_by_xpath('//*[@id="data-print-book-categories-button-proto-announce"]').click()
    time.sleep(0.5)
    # # categories

   ###  select categorie
    categries=[sheet["N"+str(i)].value,sheet["O"+str(i)].value]


    for m in range(2):
        listCat0=categries[m].split(">")
        listCat = list(map(lambda x: x.strip().lower(),listCat0))
        sizeList=len(listCat)-1 # minus 1 because the first action always happen

        # we have to lower all cat and subcat



        #this action happens always
        element = wait.until(expected_conditions.element_to_be_clickable((By.XPATH, '//*[@id="div-{0}"]/span/a'.format(listCat[0]))))
        webdriver.ActionChains(driver).move_to_element(element).click(element).perform()

        # suitch depending the number of subcats not taking into account the main cat
        if sizeList==1:
            subCat = wait.until(expected_conditions.element_to_be_clickable((By.XPATH, '//*[@id="checkbox-{0}_{1}"]'.format(listCat[0], listCat[1]))))
            subCat.click()
            webdriver.ActionChains(driver).move_to_element(element).click(element).perform()


        elif sizeList==2:

            element2 = wait.until(expected_conditions.element_to_be_clickable((By.XPATH, '//*[@id="div-{0}"]/span/a'.format(listCat[1]))))
            webdriver.ActionChains(driver).move_to_element(element2).click(element2).perform()
            subCat = wait.until(expected_conditions.element_to_be_clickable((By.XPATH, '//*[@id="checkbox-{0}_{1}"]'.format(listCat[1],listCat[2]))))
            subCat.click()
            webdriver.ActionChains(driver).move_to_element(element2).click(element2).perform()
            webdriver.ActionChains(driver).move_to_element(element).click(element).perform()

        else:

             element2 = wait.until(expected_conditions.element_to_be_clickable((By.XPATH, '//*[@id="div-{0}"]/span/a'.format(listCat[1]))))
             webdriver.ActionChains(driver).move_to_element(element2).click(element2).perform()
             element3 = wait.until(expected_conditions.element_to_be_clickable((By.XPATH, '//*[@id="div-{0}_{1}"]/span/a'.format(listCat[1],listCat[2]))))
             webdriver.ActionChains(driver).move_to_element(element2).click(element3).perform()
             subCat = wait.until(expected_conditions.element_to_be_clickable((By.XPATH, '//*[@id="checkbox-{0}_{1}_{2}"]'.format(listCat[1], listCat[2],listCat[3]))))
             subCat.click()
             webdriver.ActionChains(driver).move_to_element(element2).click(element3).perform()
             webdriver.ActionChains(driver).move_to_element(element2).click(element2).perform()
             webdriver.ActionChains(driver).move_to_element(element).click(element).perform()




    okCat = wait.until( expected_conditions.element_to_be_clickable((By.XPATH, '//*[@id="category-chooser-ok-button"]/span/input')))
    okCat.click()


   ###



    # element = wait.until(expected_conditions.element_to_be_clickable((By.XPATH, '//*[@id="div-fiction"]/span/a')))
    # webdriver.ActionChains(driver).move_to_element(element).click(element).perform()
    # subCat = wait.until(expected_conditions.element_to_be_clickable((By.XPATH, '//*[@id="checkbox-fiction_general"]')))
    # subCat.click()
    driver.maximize_window();

    # adult content
    time.sleep(1)
    # isadult=2 if sheet["P"+str(i)].value else 1;
    isAdultSelect=2 if isAdult=="Yes Only Adult" else 1;
    driver.find_element_by_xpath('//*[@id="data-print-book-is-adult-content"]/div/div/div[{}]/div/label/input'.format(isAdultSelect)).click()
    time.sleep(1)
    # save btn
    saveContinue = wait.until(expected_conditions.element_to_be_clickable((By.XPATH, '//*[@id="save-and-continue-announce"]')))
    saveContinue.click()
    time.sleep(1)

    ########################## second Page ####################3




    expandCoverUploadSelector="#data-print-book-publisher-cover-choice-accordion > div.a-box.a-last > div > div.a-accordion-row-a11y > a"
    wait.until(expected_conditions.element_to_be_clickable((By.CSS_SELECTOR,expandCoverUploadSelector))).click()
    # isnb
    Isbn=wait.until(expected_conditions.element_to_be_clickable((By.ID,"free-print-isbn-btn-announce")))
    Isbn.click();
    time.sleep(1)

    confirmIsbn=wait.until(expected_conditions.element_to_be_clickable((By.ID,"print-isbn-confirm-button-announce")))
    confirmIsbn.click();
    ##################################################3
    time.sleep(3)
    driver.maximize_window()
    NocolorPrintId="a-autoid-1-announce"
    colorPrintId="a-autoid-2-announce";
    # actualColor=colorPrintId if sheet["Q"+str(i)].value else NocolorPrintId
    actualColorSelect = colorPrintId if isColor=="Color" else NocolorPrintId
    driver.find_element_by_id(actualColorSelect).click();

    ########################3 trime size ###############

    size85x11Id = "trim-size-standard-option-1-3-announce"
    size6x9Id = "trim-size-popular-option-0-3-announce"
    # actualSize=size85x11Id if sheet["R"+str(i)].value else size6x9Id
    actualSizeSelect = size85x11Id if isSize == "8.5x11" else size6x9Id

    time.sleep(1.7)

    if actualSizeSelect==size85x11Id: # only chose all size if the size is 8.5 , else skip
        sizeId="trim-size-btn-announce"
        driver.find_element_by_id(sizeId).click();
        time.sleep(2)
        choseSize=wait.until(expected_conditions.element_to_be_clickable((By.ID,actualSizeSelect)))
        choseSize.click()

    time.sleep(4)

    #########################  Bleed or not #####################3

    NobleedId="a-autoid-3-announce"

    # so if 0 + 3 =3 whis the nobeedidisBleedExcel=4 if sheet["S"+str(i)].value == 1 else 3
    #
    isBleedSelect=4 if isBleed=="Bleed" else 3

    BleedId="a-autoid-{}-announce".format(isBleedSelect)


    bleed=wait.until(expected_conditions.element_to_be_clickable((By.ID,BleedId)))
    webdriver.ActionChains(driver).move_to_element(bleed).click(bleed).perform()
    time.sleep(1.5)

    uploadInteriorId="data-print-book-publisher-interior-file-upload-AjaxInput"
    uploadCoverId="data-print-book-publisher-cover-file-upload-AjaxInput"




    # pathToInterior="C:\\Users\\OnceLearner\\Desktop\\Kdp\\cursing\\book2\\manus.pdf"
    # pathToCover="C:\\Users\\OnceLearner\\Desktop\\Kdp\\cursing\\book2\\cover.pdf"

    driver.find_element_by_id(uploadInteriorId).send_keys(sheet["U"+str(i)].value)

    time.sleep(3)

    # upcoverFinishManus.send_keys(pathToCover)

    driver.find_element_by_id(uploadCoverId).send_keys(sheet["V"+str(i)].value);

    launchPreviewId="print-preview-announce"
    launchPreviewContinueId = "print-preview-confirm-button-announce";



    launchPreview=wait.until(expected_conditions.element_to_be_clickable((By.ID,launchPreviewId)))
    webdriver.ActionChains(driver).move_to_element(launchPreview).click(launchPreview).perform()

    time.sleep(1)

    confirm=driver.find_elements_by_id(launchPreviewContinueId)[1]
    webdriver.ActionChains(driver).move_to_element(confirm).click(confirm).perform()


   ##### confirmation
    waitalot=WebDriverWait(driver,700)


    time.sleep(1.7)
    approve="printpreview_approve_button_enabled"

    approvebtn=waitalot.until(expected_conditions.element_to_be_clickable((By.ID,approve)))
    webdriver.ActionChains(driver).move_to_element(approvebtn).click(approvebtn).perform()

    ####### return back to 2 page#####

    saveAndContinueId="save-and-continue-announce"
    saveAndContinue=waitalot.until(expected_conditions.element_to_be_clickable((By.ID,saveAndContinueId)))
    webdriver.ActionChains(driver).move_to_element(saveAndContinue).click(saveAndContinue).perform()

############## page 3 ##############


    time.sleep(1)
    priceXpath='//*[@id="data-pricing-print-us-price-input"]/input'
    waitalot.until(expected_conditions.presence_of_element_located((By.XPATH,priceXpath))).send_keys(str(sheet["T"+str(i)].value))

    publishId="save-and-publish-announce"
    saveDraftId="save-announce"
    savedraft=waitalot.until(expected_conditions.element_to_be_clickable((By.ID,saveDraftId)))
    webdriver.ActionChains(driver).move_to_element(savedraft).click(savedraft).perform()

    time.sleep(2)

    # driver.close()



# WebElement fileInput = driver.findElement(By.id("uploadFile"));
# fileInput.sendKeys("/path/to/file.jpg");


# links = ['https://www.yahoo.com', 'http://bings.com']
# for link in links:
#     control_string = "window.open('{0}')".format(link)
#     driver.execute_script(control_string)
