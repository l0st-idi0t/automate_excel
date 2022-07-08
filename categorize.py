from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementNotInteractableException, ElementClickInterceptedException

from openpyxl import load_workbook
from automate import changeActive


def main():
	wb_file = 'Customer General Code #2  6-28-2022.xlsx'

	wb = load_workbook(wb_file, data_only=True)
	changeActive('Sheet2', wb)

	sheet = wb.active

	driver = webdriver.Chrome("chromedriver.exe")

	for i in range(2, 2528):

		driver.get(f'https://maps.google.com/?q={sheet[f"F{i}"].value} {sheet[f"E{i}"].value}')

		try:
			print(driver.find_element("xpath", '/html/body/div[3]/div[9]/div[9]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/div[1]/div[1]/div[1]/h1/span[1]').text)
		except NoSuchElementException :
			print('failed')


if __name__ == "__main__":
	main()

	#//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/div[1]/div[1]/div[1]/h1/span[1]
