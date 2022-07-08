from selenium import webdriver

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementNotInteractableException, ElementClickInterceptedException

from openpyxl import load_workbook
from automate import changeActive


def main():
	wb_file = 'Customer General Code #2  6-28-2022.xlsx'

	wb = load_workbook(wb_file, data_only=True)
	changeActive('Sheet2', wb)

	sheet = wb.active

	driver = webdriver.Chrome("chromedriver.exe")

	driver.get(f'https://maps.google.com/?q={sheet["F2"].value} {sheet["E2"].value}')



if __name__ == "__main__":
	main()