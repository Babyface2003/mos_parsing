from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import pandas as pd
import os
import shutil
import glob
import time
def parse_CADDOP(driver):
    download_path = "D:\Work\Save_mos_ru"
    op = Options()
    op.add_argument('--disable-notifications')
    op.add_experimental_option("prefs",{
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    # Initializing the Chrome webdriver with the options
    driver = webdriver.Chrome(options=op)

    # Setting Chrome to trust downloads
    driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
    params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': download_path}}
    command_result = driver.execute("send_command", params)

    driver.get("https://controlpp.mos.ru/processor/back-office/login.faces")

    #вводим логин

    login_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "j_username")))
    login_field.send_keys("...")

    # вводим пароль

    password_field = driver.find_element(By.ID, "j_password")
    password_field.send_keys("...")
    password_field.send_keys(Keys.RETURN)

    WebDriverWait(driver, 10).until(EC.url_matches("https://controlpp.mos.ru/processor/back-office/index.faces"))


    dropdown_element = driver.find_element(By.ID, "mainMenuSubView:mainMenuForm:reportOnlineGroupMenu:hdr")
    dropdown_element.click()

    dropdown_menu = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "mainMenuSubView:mainMenuForm:salesReportGroupMenu")))
    dropdown_items = dropdown_menu.find_elements(By.XPATH, ".//li")

    for item in dropdown_items:
        if item.text == "Отчеты по продажам":
            item.click()
            break

    sales_report_dropdown = driver.find_element(By.ID, "mainMenuSubView:mainMenuForm:salesReportGroupMenu:hdr")
    sales_report_dropdown.click()

    time.sleep(1)  # Подождем некоторое время для загрузки выпадающего списка

    second_dropdown = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "mainMenuSubView:mainMenuForm:totalSalesReportMenuItem"))
    )

    # Ждем некоторое время после открытия второго выпадающего списка


    # Подождать, пока новый выпадающий список загрузится
    WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, "//div[@id='mainMenuSubView:mainMenuForm:salesReportGroupMenu:cnt']"))
    )

    # Найти и нажать на кнопку "Сводный отчет по продажам"
    summary_sales_report_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//div[@id='mainMenuSubView:mainMenuForm:totalSalesReportMenuItem']"))
    )
    summary_sales_report_button.click()

    #Открываем где три точки
    three_point = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH,"/html/body/table/tbody/tr[2]/td[2]/form/div/div[2]/span/table/"
                                                 "tbody/tr/td/table/tbody/tr[1]/td/div/table/tbody/tr[1]/td[2]/span/input[2]"))
    )
    three_point.click()

    #Выбираем Верону
    verona = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[56]/div[3]/div/form/table[3]/tbody/tr/td/table/tbody[1]/tr[2]/td[2]')))
    verona.click()

    #Нажимаем ОК
    OK = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[56]/div[3]/div/form/table[4]/tbody/tr/td[1]/input')))
    OK.click()
    time.sleep(1)

    #выбираем округ
    select_district = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr[2]/td[2]/form/div/div[2]/span/table/tbody/tr/td/table/tbody/tr[1]/td/div'
                                                  '/table/tbody/tr[2]/td[2]/span/input')))
    select_district.click()
    time.sleep(1)

    #Выбираем школы исключения, которые етсь на цао

    #Нажимаем на ещё одну окей

    select_mok = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[56]/div[3]/div/form/table/tbody/tr[1]/td/table[1]/tbody/tr/td[1]'
                                                  '/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td[2]/input')))
    select_mok.click()
    time.sleep(1)
    select_mok.send_keys('"ГБПОУ ""1-й МОК""-12"')
    select_mok.send_keys(Keys.RETURN)
    select = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[56]/div[3]/div/form/table/tbody/tr[1]/td/table[2]/tbody/tr/td[1]/input')))
    select.click()
    time.sleep(2)
    select_mok.clear()

    select_mgpu = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[56]/div[3]/div/form/table/tbody/tr[1]/td/table[1]/tbody/tr/td[1]'
                       '/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td[2]/input')))
    select_mgpu.click()
    time.sleep(2)
    select_mgpu.send_keys('ГБОУ ВПО МГПУ-3')
    select_mgpu.send_keys(Keys.RETURN)
    select = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[56]/div[3]/div/form/table/tbody/tr[1]/td/table[2]/tbody/tr/td[1]/input')))
    select.click()
    time.sleep(2)
    select_mgpu.clear()
    time.sleep(2)
    select_kgt = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[56]/div[3]/div/form/table/tbody/tr[1]/td/table[1]/tbody/tr/td[1]'
                       '/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td[2]/input')))
    select_kgt.click()
    time.sleep(2)
    select_kgt.send_keys('ГБПОУ КЖГТ-')
    select_kgt.send_keys(Keys.RETURN)
    time.sleep(1)
    select = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[56]/div[3]/div/form/table/tbody/tr[1]/td/table[2]/tbody/tr/td[1]/input')))
    select.click()
    time.sleep(2)
    select_kgt.clear()

    select_kc = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[56]/div[3]/div/form/table/tbody/tr[1]/td/table[1]/tbody/tr/td[1]'
                       '/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td[2]/input')))
    select_kc.click()
    time.sleep(2)
    select_kc.send_keys('ГБПОУ КС № 54-2023')
    select_kc.send_keys(Keys.RETURN)
    select = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[56]/div[3]/div/form/table/tbody/tr[1]/td/table[2]/tbody/tr/td[1]/input')))
    select.click()
    time.sleep(2)
    select_kc.clear()

    time.sleep(2)

    ok = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[56]/div[3]/div/form/table/tbody/tr[3]/td/span/input[1]')))
    ok.click()
    time.sleep(2)
    # выбираем один день

    select_data = WebDriverWait(driver,10).until(
        EC.presence_of_element_located((By.XPATH,'/html/body/table/tbody/tr[2]/td[2]/form/div/div[2]/span/table/tbody'
                                                 '/tr/td/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/select')))
    select_data.click()

    time.sleep(1)
    select_data_to_do = Select(select_data)
    select_data_to_do.select_by_value('ONE_DAY')

    #теперь скачиваем excel файл

    download_fail = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr[2]/td[2]/form/div/div[2]/span/table/tbody/tr/td/table/tbody/tr[2]/td'
                                                  '/table/tbody/tr/td[2]/input')))
    download_fail.click()

    time.sleep(1)


    # Находим все файлы с расширением .xls в директории загрузки
    xls_files = glob.glob(os.path.join(download_path, "*.xls"))

    # Проверяем, что список файлов не пустой
    if xls_files:
        # Выбираем последний среди найденных файлов
        latest_xls_file = max(xls_files, key=os.path.getctime)

        # Переименовываем последний файл в желаемое имя
        new_filename = os.path.join(download_path, "ЦАОДОП.xls")  # Замените на желаемое новое имя файла
        shutil.move(latest_xls_file, new_filename)

    time.sleep(2)
    # Возвращаем данные
    return pd.DataFrame({"district": ["ZAO", "ZELAO", "CAO", "CZAO", "CVAO", "CAD", "CADDOP", "UAO", "UVAO", "UZAO"]})

