from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementNotInteractableException, ElementClickInterceptedException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

from openpyxl import load_workbook
from automate import changeActive
from fuzzywuzzy import fuzz


# F	Chinese Buffet
# G	Chinese Dine-in
# H	Chinese Take-out
# I	Cajun - Seafood
# J	火锅 Hot Pot
# K	烧烤 BBQ
# L	Japanese
# M	American
# N	东南亚 （SE Asian）
# O	Other	
# Q	Mexican
# R	Indian Buffet
# T	CAN'T FOUND
# U	NOT RESTAURANT
# V	closed
# W	Wholesale


def main():
	wb_file = 'Customer General Code #2  6-28-2022.xlsx'

	wb = load_workbook(wb_file, data_only=True)
	changeActive('Sheet2', wb)

	sheet = wb.active

	driver = webdriver.Chrome("chromedriver.exe")

	for i in range(2, 50):
		address = sheet[f"E{i}"].value
		name = sheet[f"F{i}"].value.replace("&", "and")

		driver.get(f'https://maps.google.com/?q={name} {address}')

		title = "/html/body/div[3]/div[9]/div[9]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/div[1]/div[1]/div[1]/h1/span[1]"
		description = "PYvSYb"
		services = "E0DTEd"
		address = "Io6YTe fontBodyMedium" 
		type_restaurant = "DkEaL"
		reviews = "m6QErb tLjsW"  
		
		def checks():
			types = dict.fromkeys(['thai', 'viet', 'india'], 'N')
			types.update(dict.fromkeys(['buffet'], 'F'))
			types.update(dict.fromkeys(['cajun', 'seafood'], 'I'))
			types.update(dict.fromkeys(['wholesale'], 'W'))
			types.update(dict.fromkeys(['sushi', 'japan'], 'L'))
			types.update(dict.fromkeys(['barbeque', 'bbq'], 'K'))
			types.update(dict.fromkeys(['hot pot'], 'J'))


			print(f'{i}: {driver.find_element("xpath", title).text}')
			if "closed" in driver.find_element(By.CLASS_NAME, type_restaurant).text.lower():
				sheet[f"H{i}"] = "V"
				return

			for i in ["wholesale", "thai", "viet", "india", "sushi", "japan", "barbeque", "bbq", "hot pot", "cajun", "seafood", "buffet"]:
				if i in driver.find_element(By.CLASS_NAME, type_restaurant).text.lower() or i in driver.find_element("xpath", title).text.lower():
					sheet[f"H{i}"] = types[i]
					return

			


		try:
			checks()

		except NoSuchElementException:
			try:
				WebDriverWait(driver, timeout=3).until(lambda d: d.find_element(By.CLASS_NAME, "hfpxzc"))
				driver.find_element(By.CLASS_NAME, "hfpxzc").click()
				WebDriverWait(driver, timeout=3).until(lambda d: d.find_element("xpath", title))
				print(f'{i}: {driver.find_element("xpath", title).text}')
			except NoSuchElementException:
				print(i)
			



			# result_names = [i.text for i in driver.find_elements(By.CSS_SELECTOR, 'div[class="NrDZNb"]')]
			# print(f'{i}: {[(fuzz.token_sort_ratio(name, k), fuzz.partial_ratio(name.lower(), k.lower())) for k in result_names]}')
			# print(result_names)



if __name__ == "__main__":
	main()



