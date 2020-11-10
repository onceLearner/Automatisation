import openpyxl
# from selenium import webdriver
# from selenium.webdriver import ActionChains
# from selenium.webdriver.common.keys import Keys
#
#
workbook=openpyxl.load_workbook("C:\\Users\\OnceLearner\\Desktop\\Kdp\\notebook\\coWorker\\data.xlsx",data_only=True)
sheet=workbook.active
i=22
listCat0=sheet["N"+str(i)].value.split(">")
sizeList=len(listCat0)-1 # minus 1 because the first action always happen

        # we have to lower all cat and subcat
listCat = list(map(lambda x: x.strip().lower(),listCat0))
print('//*[@id="div-{0}_{1}"]/span/a'.format(listCat[1],listCat[2])))

#
# driver = webdriver.Chrome()
# driver.get("http://www.google.com/")
#
# #open tab
# driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL +"t")
# You can use (Keys.CONTROL + 't') on other OSs

# Load a page
# driver.get('http://stackoverflow.com/')


# tab1=driver.window_handles[1]
# driver.switch_to.window(tab1)
# driver.get("https://gmail.com")


from selenium import webdriver
from selenium.webdriver.common.keys import Keys

# driver = webdriver.Chrome()
# driver.get("https://www.google.com/")
# # links = ['https://twitter.com', 'link_2', 'link_3']
# # driver.execute_script("window.open('{}')".format(links[0]))
# # driver.switch_to.window(driver.window_handles[0])
#
# from joblib import Parallel, delayed
# urls=["https://goole.com","https://twitter.com"]
#
# def do_stuff(url):
#     driver.get(url);
#
#
#
#
# Parallel(n_jobs=-1)(delayed(do_stuff)(url) for url in urls) #







