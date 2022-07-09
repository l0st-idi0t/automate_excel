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

	i = 72

	title = "/html/body/div[3]/div[9]/div[9]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/div[1]/div[1]/div[1]/h1/span[1]"
	description = "PYvSYb"
	services = "E0DTEd"
	r_address = "rogA2c" 
	type_restaurant = "button[jsan='7.DkEaL,0.jsaction']"
	reviews = "div[role='radiogroup']"

	def checks(address, name):
			types = dict.fromkeys(['thai', 'viet', 'india'], 'N')
			types.update(dict.fromkeys(['buffet'], 'F'))
			types.update(dict.fromkeys(['cajun', 'seafood'], 'I'))
			types.update(dict.fromkeys(['wholesale'], 'W'))
			types.update(dict.fromkeys(['sushi', 'japan', 'tokyo'], 'L'))
			types.update(dict.fromkeys(['barbeque', 'bbq'], 'K'))
			types.update(dict.fromkeys(['hot pot'], 'J'))
			types.update(dict.fromkeys(['mexic'], 'Q'))
			types.update(dict.fromkeys(['dine-in'], 'G'))
			types.update(dict.fromkeys(['takeout, drive-through'], 'G'))




			try:
				driver.find_element(By.CSS_SELECTOR, 'span[style="color:#D93025"]')
				sheet[f"H{i}"] = "V"
				print(f'{i}: {driver.find_element("xpath", title).text} and type is closed')
				return 
			except Exception as e:
				pass


			WebDriverWait(driver, timeout=1).until(lambda d: d.find_element(By.CSS_SELECTOR, type_restaurant))
			restaurant_type = driver.find_element(By.CSS_SELECTOR, type_restaurant).text
			restaurant_title = driver.find_element("xpath", title).text
			WebDriverWait(driver, timeout=1).until(lambda d: d.find_element(By.CLASS_NAME, r_address))
			restaurant_addr = driver.find_element(By.CLASS_NAME, r_address).text
			restaurant_reviews = []

			try:
				restaurant_reviews = driver.find_element(By.CSS_SELECTOR, reviews).lower().split('\n')
				print(restaurant_reviews)
			except Exception as e:
				pass

			print(f'{i}: {restaurant_title} and type is {restaurant_type}')


			one = address.replace('-', '').replace('_', '') if not sheet[f'E{i}'].value.split('-')[-1].isdigit() else sheet[f'E{i}'].value[:-4].replace('-', '').replace('_', '')
			two = restaurant_addr

			onearr = ''.join([k for k in [char for char in one] if k.isdigit()])
			twoarr = ''.join([k for k in [char for char in two] if k.isdigit()])


			if (fuzz.token_sort_ratio(one, two) < 70 and fuzz.token_sort_ratio(onearr, twoarr) < 85 and fuzz.partial_ratio(one.lower(), two.lower()) < 90):
				sheet[f"H{i}"] = "T"
				return


			try:
				for keyword in ["wholesale", "thai", "viet", "india", "mexic", "sushi", "japan", "tokyo", "barbeque", "bbq", "hot pot", "cajun", "seafood", "buffet"]:
					if keyword in restaurant_type.lower() or keyword in restaurant_title.lower() or len([s for s in restaurant_reviews if keyword in s]) != 0:
						sheet[f"H{i}"] = types[keyword]
						return
			except Exception as e:
				pass


			try:
				WebDriverWait(driver, timeout=1).until(lambda d: d.find_element(By.CLASS_NAME, services))
				sheet[f"H{i}"] = types[driver.find_element(By.CLASS_NAME, services).text.lower().split('\n')[0]]
			except Exception as e:
				sheet[f"H{i}"] = "T"


			return


	while i <= 80:
		address = sheet[f"E{i}"].value
		name = sheet[f"F{i}"].value.replace("&", "and").replace("#", "")

		driver.get(f'https://maps.google.com/?q={name} {address}')

		try:
			checks(address, name)
		except Exception as e:
			try:
				WebDriverWait(driver, timeout=2).until(lambda d: d.find_element(By.CLASS_NAME, "hfpxzc"))
				driver.find_element(By.CLASS_NAME, "hfpxzc").click()
				WebDriverWait(driver, timeout=2).until(lambda d: d.find_element("xpath", title))
				checks(address, name)
			except Exception as e:
				print(f'{i}: error')
		
		i += 1



			# result_names = [i.text for i in driver.find_elements(By.CSS_SELECTOR, 'div[class="NrDZNb"]')]
			# print(f'{i}: {[(fuzz.token_sort_ratio(name, k), fuzz.partial_ratio(name.lower(), k.lower())) for k in result_names]}')
			# print(result_names)

	wb.save(filename = wb_file)


if __name__ == "__main__":
	main()



