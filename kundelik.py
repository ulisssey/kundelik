from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.remote.remote_connection import LOGGER
import pandas as pd
import time
import sys
import re
import logging


not_found_subjects = []
not_found_teachers = []
not_found_subgroups = []
chrome_options = webdriver.ChromeOptions()
prefs = {"profile.default_content_setting_values.notifications": 2}
chrome_options.add_experimental_option("prefs", prefs)

def fill(Item_name, Teacher_name, driver):
    subgroup_name = ''
    if '(' in Item_name:
        splits = Item_name.split('(')
        Item_name = splits[0].strip()
        subgroup_name = splits[1][:-1]
    select = Select(driver.find_element(By.ID, 'subject'))
    try:
        select.select_by_visible_text(Item_name)
    except:
        not_found_subjects.append(Item_name) 
        return 0      
    time.sleep(1)

    all_teacher = WebDriverWait(driver, 10).until(lambda driver: driver.find_elements(By.XPATH, '//select[@id="teacher"]/option'))
    if len(Teacher_name) < 2:
        not_found_teachers.append(Teacher_name)
    elif 'Нет учителя' in Teacher_name:
        pass
    else:
        try:
            for x in all_teacher:
                if x.text.split(' ')[0] == Teacher_name.split(' ')[0]:
                    Teacher_name = x.text
                    break
        except:
            pass

        WebDriverWait(driver, 60).until(lambda driver: driver.find_element(By.XPATH, "//select[@id='teacher']")).click()
        select_element = driver.find_element(By.XPATH, "//select[@id='teacher']")
        select = Select(select_element)
        try:
            select.select_by_visible_text(Teacher_name)
        except:
            not_found_teachers.append(Teacher_name)
        time.sleep(0.5)

    if subgroup_name != '':
        all_subgroups = WebDriverWait(driver, 10).until(lambda driver: driver.find_elements(By.XPATH, "//select[@id='subgroup']/option"))
        x = 0
        try:
            for s in all_subgroups:
                if subgroup_name in s.text:
                    subgroup_name = s.text
                    x = 1
                    break
        except:
            not_found_subgroups.append(subgroup_name)
        WebDriverWait(driver, 60).until(lambda driver: driver.find_element(By.XPATH, "//select[@id='subgroup']")).click()
        select = Select(driver.find_element(By.ID, 'subgroup'))
        try:
            select.select_by_visible_text(subgroup_name)
        except:
            not_found_subgroups.append(subgroup_name)
            return 0
        time.sleep(0.5)

def fill_schedule(login, password, file):
    LOGGER.setLevel(logging.WARNING)
    driver = webdriver.Chrome(options=chrome_options, service=Service(executable_path=executable_path))

    driver.get('https://schools.kundelik.kz/schedules/')
    time.sleep(1)

    inputElement1 = driver.find_element(By.NAME, "login")
    inputElement1.send_keys(login)

    inputElement2 = driver.find_element(By.NAME, "password")
    inputElement2.send_keys(password)

    inputElement2.send_keys(Keys.ENTER)

    try:
        if WebDriverWait(driver, 5).until(lambda driver: driver.find_element(By.XPATH, "//div[contains(text(), 'Пайдаланушы аты немесе құпиясөзде қате бар. Өрістердің дұрыс толтырылуын тексеріңіз.')]")):
            driver.quit()
            return "1"
    except:
        pass
    driver.maximize_window()
   
    df = pd.read_excel(file)

    a = list(df.columns)
    class_name = re.split(r'Класс: ', a[0])[1]
    try:
        driver.find_element(By.XPATH, "//*[text()='%s']" % class_name.strip()).click()
    except:
        driver.quit()
        return "20"
    driver.find_element(By.CSS_SELECTOR, '[id="linktextDataGenerator"]').click()

    quarter_num = re.split(r' четверть, ', df[a[0]][0])[0]
    driver.find_element(By.XPATH, "//a[@id='buttonCreateSchedule']").click()
    WebDriverWait(driver, 60).until(lambda driver: driver.find_element(By.XPATH, "//input[@id='name']")).send_keys('Расписание на %s четверть' % quarter_num)
    driver.find_element(By.XPATH, "//input[@id='save']").click()
    WebDriverWait(driver, 60).until(lambda driver: driver.find_element(By.XPATH, "//div[contains(text(), 'успешно создана')]"))

    index_df_1 = df[(df['Unnamed: 1'] == 1)].index
    i = 0
    n1 = (index_df_1[i] - 1)
    if i != len(index_df_1) - 1:
        n2 = (index_df_1[i + 1] - 1)
    else:
        n2 = len(df)

    df_data = (df.iloc[n1:n2]).reset_index(drop=True)

    n = list(df_data[(df_data['Unnamed: 1'] != ' ') & (df_data['Unnamed: 1'].notnull())]['Unnamed: 1'])
    if 0 in n:
    # try:
        for i in range(len(n) - 1):
            x = ''
            s = 0
            if df_data['Unnamed: 1'][i + 1] == '#':
                df_test = df_data[i :i + 2].reset_index(drop=True)
                x = 1
            if df_data['Unnamed: 1'][i ] != '#' and type(df_data['Unnamed: 1'][i + 1]) == int:
                df_test = df_data[i :i + 1].reset_index(drop=True)
                x = 1
            tr_index = df_test['Unnamed: 1'][0] + 1
            if x != '':
                for j in range(5):
                    if df_test[a[j + 2]][0] == 'Праздничный день':
                        pass
                    elif df_test[a[j + 2]][0] != ' ':
                        d = driver.find_element(By.CSS_SELECTOR, '[class="grid scheduleEditor clear"]')
                        to_element = d.find_elements(By.CSS_SELECTOR, 'tr')[tr_index].find_elements(By.CSS_SELECTOR, 'td')[j + 1]
                        driver.execute_script("arguments[0].scrollIntoView();", to_element)
                        to_element.click()
                        time.sleep(0.5)

                        for k in range(len(df_test)):
                            if df_test[a[j + 2]][k] != ' ':
                                Item_name = re.split(r'\n', df_test[a[j + 2]][k])[0]
                                Teacher_name = re.split(r'\n', df_test[a[j + 2]][k])[1]

                                if len(df_test) == 2:
                                    if df_test[a[j + 2]][1] == ' ':
                                        status = fill(Item_name, Teacher_name, driver)
                                        if status == 0:
                                            driver.find_element(By.XPATH, "//a[text()='Отмена']").click()
                                        else:
                                            driver.find_element(By.XPATH, "//a[text()='Создать']").click()
                                        time.sleep(1)

                                    if df_test[a[j + 2]][1] != ' ' and k == 0:
                                        status = fill(Item_name, Teacher_name, driver)
                                        if status == 0:
                                            driver.find_element(By.XPATH, "//a[text()='Отмена']").click()
                                            s = 0
                                        else:
                                            driver.find_element(By.XPATH, "//a[text()='Создать и добавить ещё']").click()
                                            s = 1
                                        time.sleep(1)

                                    if df_test[a[j + 2]][1] != ' ' and k == 1:
                                        if s == 0:
                                            pass
                                        else:
                                            status = fill(Item_name, Teacher_name, driver)
                                            if status == 0:
                                                driver.find_element(By.XPATH, "//a[text()='Отмена']").click()
                                            else:
                                                driver.find_element(By.XPATH, "//a[text()='Создать']").click()
                                            time.sleep(1)
                                else:
                                    status = fill(Item_name, Teacher_name, driver)
                                    if status == 0:
                                        driver.find_element(By.XPATH, "//a[text()='Отмена']").click()
                                    else:
                                        driver.find_element(By.XPATH, "//a[text()='Создать']").click()
                                    time.sleep(1)
                            else:
                                pass
    else:
        for i in range(len(n) - 1):
            x = ''
            s = 0
            if df_data['Unnamed: 1'][i + 2] == '#':
                df_test = df_data[i + 1:i + 3].reset_index(drop=True)
                x = 1
            if df_data['Unnamed: 1'][i + 1] != '#' and type(df_data['Unnamed: 1'][i + 2]) == int:
                df_test = df_data[i + 1:i + 2].reset_index(drop=True)
                x = 1
            tr_index = df_test['Unnamed: 1'][0]
            if x != '':
                for j in range(5):
                    if df_test[a[j + 2]][0] == 'Праздничный день':
                        pass
                    elif df_test[a[j + 2]][0] != ' ':
                        d = driver.find_element(By.CSS_SELECTOR, '[class="grid scheduleEditor clear"]')
                        to_element = d.find_elements(By.CSS_SELECTOR, 'tr')[tr_index].find_elements(By.CSS_SELECTOR, 'td')[j + 1]
                        driver.execute_script("arguments[0].scrollIntoView();", to_element)
                        to_element.click()
                        time.sleep(0.5)

                        for k in range(len(df_test)):
                            if df_test[a[j + 2]][k] != ' ':
                                Item_name = re.split(r'\n', df_test[a[j + 2]][k])[0]
                                Teacher_name = re.split(r'\n', df_test[a[j + 2]][k])[1]

                                if len(df_test) == 2:
                                    if df_test[a[j + 2]][1] == ' ':
                                        status = fill(Item_name, Teacher_name, driver)
                                        if status == 0:
                                            driver.find_element(By.XPATH, "//a[text()='Отмена']").click()
                                        else:
                                            driver.find_element(By.XPATH, "//a[text()='Создать']").click()
                                        time.sleep(1)

                                    if df_test[a[j + 2]][1] != ' ' and k == 0:
                                        status = fill(Item_name, Teacher_name, driver)
                                        if status == 0:
                                            driver.find_element(By.XPATH, "//a[text()='Отмена']").click()
                                            s = 0
                                        else:
                                            driver.find_element(By.XPATH, "//a[text()='Создать и добавить ещё']").click()
                                            s = 1
                                        time.sleep(1)

                                    if df_test[a[j + 2]][1] != ' ' and k == 1:
                                        if s == 0:
                                            pass
                                        else:
                                            status = fill(Item_name, Teacher_name, driver)
                                            if status == 0:
                                                driver.find_element(By.XPATH, "//a[text()='Отмена']").click()
                                            else:
                                                driver.find_element(By.XPATH, "//a[text()='Создать']").click()
                                            time.sleep(1)
                                else:
                                    status = fill(Item_name, Teacher_name, driver)
                                    if status == 0:
                                        driver.find_element(By.XPATH, "//a[text()='Отмена']").click()
                                    else:
                                        driver.find_element(By.XPATH, "//a[text()='Создать']").click()
                                    time.sleep(1)
                            else:
                                pass
    # except:
    #     return "4"
    driver.quit()
    all_not_found = []
    if not_found_subjects != []:
        return "15"
    elif not_found_teachers != []:
        print(not_found_teachers)
        return "23"
    elif not_found_subgroups != []:
        return "24"
    else:
        return "0"

    # return not_found_subjects, not_found_teachers, not_found_subgroups
import argparse
parser = argparse.ArgumentParser()
parser.add_argument('--login', type=str, required=True)
parser.add_argument('--password', type=str, required=True)
parser.add_argument('--file', type=str, required=True)
parser.add_argument('--chromedriver', type=str)
parser.add_argument('--chromium', type=str)
args = parser.parse_args()
executable_path = args.chromedriver
chrome_options.binary_location = args.chromium
result = fill_schedule(args.login, args.password, args.file)
sys.stdout.write(result)
