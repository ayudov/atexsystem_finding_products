import xlrd
from openpyxl import *
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ChromeOptions

#Get file names

loc = ("redirections.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)
list_of_product_names = []
for i in range(sheet.nrows):
	list_of_product_names.append(sheet.cell_value(i, 0))
print(list_of_product_names)

# Using Chrome to access web
chrome_options = ChromeOptions()
chrome_options.add_argument("--start-maximized")
driver = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)
driver.get('http://www.atex-system.com/cgv/')
wb=load_workbook('redirections_links.xlsx')

for i in range(len(list_of_product_names)):
	try:
		elem = driver.find_element_by_class_name('icon-magnifier').click()
		elem = driver.find_element_by_class_name('field')
		elem.send_keys(list_of_product_names[i])
		elem.send_keys(Keys.ENTER)
		elem = driver.find_element_by_class_name('search-entry-title.entry-title').click()
		print(str(i) + ': ' + driver.current_url)
		ws = wb["Sheet1"]
		wcell1 = ws.cell(i+1,1)
		wcell1.value = str(driver.current_url)
	except:
		wcell1.value = str('ERROR')
		print(str(i+1) + ": Couldn't find '" + str (list_of_product_names[i] + "' on the site"))

wb.save('redirections_links.xlsx')
driver.quit()