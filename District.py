from selenium import webdriver
from ZAO import parse_ZAO
from ZELAO import parse_ZELAO
from CAO import parse_CAO
from CZAO import parse_CZAO
from CAD import parse_CAD
from CADDOP import parse_CADDOP
from CVAO import parse_CVAO
from UAO import parse_UAO
from UVAO import parse_UVAO
from UZAO import parse_UZAO
import time

def save_data(data, district):
    filename = f"{district}.xlsx"
    data.to_excel(filename, index=False)

if __name__ == "__main__":
    driver = webdriver.Chrome()

    districts = ["ЗАО",'ЗЕЛАО',"САО","СЗАО","СВАО","ЦАО","ЦАОДОП","ЮАО","ЮВАО","ЮЗАО"]
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
        elif district == 'СЗАО':
            parsed_data = parse_CZAO(driver)
            save_data(parsed_data,district)
        elif district == 'СВАО':
            parsed_data = parse_CVAO(driver)
            save_data(parsed_data,district)
        elif district == 'ЦАО':
            parsed_data = parse_CAD(driver)
            save_data(parsed_data,district)
        elif district == 'ЦАОДОП':
            parsed_data = parse_CADDOP(driver)
            save_data(parsed_data,district)
        elif district == 'ЮАО':
            parsed_data = parse_UAO(driver)
            save_data(parsed_data, district)
        elif district == 'ЮВАО':
            parsed_data = parse_UVAO(driver)
            save_data(parsed_data, district)
        elif district == 'ЮЗАО':
            parsed_data = parse_UZAO(driver)
            save_data(parsed_data, district)

    driver.quit()
