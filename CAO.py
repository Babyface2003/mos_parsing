from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import pandas as pd
import os
import glob
import shutil
import time

def parse_CAO(driver):
    download_path = "D:\Work\Save_mos_ru"
    op = Options()
    op.add_argument('--disable-notifications')
    op.add_experimental_option("prefs", {
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

    # вводим логин

    login_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "j_username")))
    login_field.send_keys("ShirokovaNI")

    # вводим пароль

    password_field = driver.find_element(By.ID, "j_password")
    password_field.send_keys("b405deGB1")
    password_field.send_keys(Keys.RETURN)

    WebDriverWait(driver, 10).until(EC.url_matches("https://controlpp.mos.ru/processor/back-office/index.faces"))

    dropdown_element = driver.find_element(By.ID, "mainMenuSubView:mainMenuForm:reportOnlineGroupMenu:hdr")
    dropdown_element.click()

    dropdown_menu = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "mainMenuSubView:mainMenuForm:salesReportGroupMenu")))
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
        EC.visibility_of_element_located(
            (By.XPATH, "//div[@id='mainMenuSubView:mainMenuForm:salesReportGroupMenu:cnt']"))
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
    rsc = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[56]/div[3]/div/form/table[3]/tbody/tr/td/table/tbody[1]/tr[3]')))
    rsc.click()

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
    #убираем кнопки, чтобы не скачалось лишнего
    dop_study_button1 = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[56]/div[3]/div/form/table/tbody/tr[1]/td/table[1]/tbody/tr/td[1]/table/tbody/tr/td/table'
                                                  '/tbody/tr[3]/td/div/table/tbody/tr/td/table[5]/tbody/tr/td[1]/input')))
    dop_study_button1.click()

    time.sleep(1)
    dop_study_button2 = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[56]/div[3]/div/form/table/tbody/tr[1]/td/table[1]/tbody/tr/td[1]/table/tbody/tr/td'
                                                  '/table/tbody/tr[3]/td/div/table/tbody/tr/td/table[2]/tbody/tr/td[1]/input')))
    dop_study_button2.click()
    time.sleep(1)
    #Выпадающий список для того, чтобы выбрать округ

    cao_district_list = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[56]/div[3]/div/form/table/tbody/tr[1]/td/table[1]/tbody/tr/td[1]/table/tbody/tr/td/table'
                                                  '/tbody/tr[2]/td/table/tbody/tr[5]/td[2]/select')))
    cao_district_list.click()
    time.sleep(1)

    select_cao = Select(cao_district_list)
    select_cao.select_by_value('САО')

    #Нажимаем на кнопку выбрать всё
    select_all = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[56]/div[3]/div/form/table/tbody/tr[1]/td'
                                                  '/table[2]/tbody/tr/td[1]/input')))
    select_all.click()

    time.sleep(1)

    #Нажимаем на ещё одну окей

    ok_all = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[56]/div[3]/div/form/table/tbody/tr[3]/td/span/input[1]')))
    ok_all.click()

    time.sleep(1)

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
        new_filename = os.path.join(download_path, "САО.xls")  # Замените на желаемое новое имя файла
        shutil.move(latest_xls_file, new_filename)

    time.sleep(2)

    return pd.DataFrame({"district": ["ZAO", "ZELAO","CAO","CZAO","CVAO","CAD","CAD_DOP","UAO","UVAO","UZAO"]})

