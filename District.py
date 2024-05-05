from selenium import webdriver
from ZAO import parse_ZAO
from ZELAO import parse_ZELAO
from CAO import parse_CAO
import time

def save_data(data, district):
    filename = f"{district}.xlsx"
    data.to_excel(filename, index=False)

if __name__ == "__main__":
    driver = webdriver.Chrome()

    districts = ["ЗАО",'ЗЕЛАО',"САО"]
    for district in districts:
        if district == 'ЗАО':
            parsed_data = parse_ZAO(driver)
            save_data(parsed_data, district)
        elif district == 'ЗЕЛАО':
            parsed_data = parse_ZELAO(driver)
            save_data(parsed_data, district)
        elif district == 'САО':
            parsed_data = parse_CAO(driver)
            save_data(parsed_data, district)

    driver.quit()
