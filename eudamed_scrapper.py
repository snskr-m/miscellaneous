from selenium import webdriver
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import re
import os
from datetime import datetime
from openpyxl import load_workbook
import sys

driver = webdriver.Chrome()
keywords=[]
limit_product=2
limit_record=27



#read keywords from excel sheet
df=pd.read_excel(f'{os.getcwd()}/keywords.xlsx')
for keyword in df['keywords']:
    if isinstance(keyword,str):
        keywords.append(keyword.strip())
# keywords = ["Laparoscope", "Trocar", "Cannula", "Instruments", "Electrosurgical", "Insufflator", "Robots", "Staplers", "Suturing", "Light", "Video", "Navigation", "Hemostatic", "Smoke", "Catheters", "Probes", "Disposable"]
print(keywords)

def check_cookies():
    try:
        cookies = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="cookie-consent-banner"]/div/div/div[2]/a[1]')))
        cookies.click()
        driver.find_element(By.XPATH, '//*[@id="cookie-consent-banner"]/div/div/div[2]/button/span/span').click()
    
    except:
        pass
    time.sleep(0)


def get_product_list(search_query):
    product_list=[]
    driver.get("https://ec.europa.eu/tools/eudamed/#/screen/search-device")
    driver.refresh()
    time.sleep(12)
    check_cookies()
    search_description = driver.find_element(By.ID, "nomenclatureCode")
    actions = ActionChains(driver)
    actions.move_to_element(search_description).click().send_keys(search_query).perform()
    time.sleep(5)
    soup= BeautifulSoup(driver.page_source, "html.parser")
    list= soup.find_all("span", class_="ng-tns-c98-13 ng-star-inserted")
    for elements in list:
        product_list.append(elements.get_text().strip())
    return len(product_list)

# <li role="option" pripple="" class="p-ripple p-element p-autocomplete-item ng-tns-c98-13 ng-star-inserted" id="" style=""><span class="ng-tns-c98-13 ng-star-inserted">M040

def searching_product(search_query,i):
    driver.get("https://ec.europa.eu/tools/eudamed/#/screen/search-device")
    driver.refresh()
    time.sleep(12)
    search_description = driver.find_element(By.ID, "nomenclatureCode")
    actions = ActionChains(driver)
    actions.move_to_element(search_description).click().send_keys(Keys.DELETE)
    time.sleep(2)
    actions.move_to_element(search_description).click().send_keys(search_query).perform()
    time.sleep(2)
    path = f'//*[@id="pr_id_2_list"]/li[{i}]'
    driver.find_element(By.XPATH, path).click()
    driver.find_element(By.XPATH, '//*[@id="cdk-accordion-child-0"]/div/form/div[10]/button[1]').click()
    time.sleep(10)

def get_record_count():
    time.sleep(3)
    soup= BeautifulSoup(driver.page_source, "html.parser")
    total_count= soup.find('h2',class_='nb-records ng-star-inserted')
    text = total_count.get_text().strip()
    match = re.search(r'\d+', text)
    count = int(match.group())
    return count
  
#get selected field data from final page
def get_data(i):
    time.sleep(15)
    path = f'/html/body/app-root/eui-block-content/div/ecl-app/div/div/div/app-search-device/div/eui-block-content/div/div/p-table/div/div/table/tbody/tr[{i}]'
    driver.find_element(By.XPATH, path).click()
    # /html/body/app-root/eui-block-content/div/ecl-app/div/div/div/app-search-device/div/eui-block-content/div/div/p-table/div/div/table/tbody/tr[1]
    time.sleep(5)
    soup = BeautifulSoup(driver.page_source, "html.parser")
    content = soup.find('div',class_='main')
    rows = content.findAll('dl',class_="row ng-star-inserted")
    data = {
    'Actor/Organisation name': 'NA',
    'Applicable legislation': 'NA',
    'Risk class': 'NA',
    'Device name': 'NA',
    'Nomenclature code(s)': 'NA',
    'Device model': 'NA',
    'Name/Trade name(s)': 'NA',
    'Status': 'NA',
    'Member State where the device is or is to be made available': 'NA',
    'Country': 'NA'}
                
    for row in rows:
        field_name = row.dt.get_text().strip()
        field_value = row.dd.get_text().strip()

        if field_name in data:
            data[field_name] = field_value

    driver.back()
    time.sleep(5)
    return data

def next_page():
    time.sleep(5)
    next_page_button= driver.find_element(By.XPATH, '/html/body/app-root/eui-block-content/div/ecl-app/div/div/div/app-search-device/div/eui-block-content/div/div/p-table/div/p-paginator/div/button[3]')
    # /html/body/app-root/eui-block-content/div/ecl-app/div/div/div/app-search-device/div/eui-block-content/div/div/p-table/div/p-paginator/div/button[3]
    next_page_button.click()
    time.sleep(20)


def main(search_query):
    try:
        product_count = get_product_list(search_query)
        time.sleep(5)
        # if product_count == 0:
        #     print(f"No products found for search query: {search_query}")
        #     return
        for i in range(1,product_count+1):
            try:
                # if i>limit_product:
                #     break
                searching_product(search_query,i)
                
                record_count = get_record_count()
                pages=record_count//25
                for page in range(1,pages+2):
                    all_data = []
                    if record_count==0:
                        print("No record founds")
                        continue
                    elif record_count<=25:
                        for item in range(1,record_count+1):
                            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                            try:
                                # if item>limit_record:
                                #     break
                                all_data.append(get_data(item))
                                
                            except Exception as e:
                                print(e)
                                continue
                        df = pd.DataFrame.from_dict(all_data)
                        df.to_excel(f'{os.getcwd()}/{search_query}_{timestamp}.xlsx')
                    else:
                        for item in range(1,26):
                            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                            try:
                                # if item>limit_record:
                                #     break
                                all_data.append(get_data(item))
                                
                            except Exception as e:
                                print(e)
                                continue
                        df = pd.DataFrame.from_dict(all_data)
                        df.to_excel(f'{os.getcwd()}/{search_query}_{timestamp}.xlsx', index=False)
                        next_page()
                        time.sleep(10)
            except Exception as e:
                # timestamp=datetime.now().strftime("%Y%m%d%H%M%S")
                # error_message = str(e)
                # file_path = (f'{os.getcwd()}/error_log_{timestamp}.txt')
                # with open(file_path, 'w') as file:
                #     file.write(error_message)
                # print(f'Error message has been stored in {file_path}')
                continue
    except Exception as e:
        # timestamp=datetime.now().strftime("%Y%m%d%H%M%S")
        # error_message = str(e)
        # file_path = (f'{os.getcwd()}/error_log2_{timestamp}.txt')
        # with open(file_path, 'w') as file:
        #     file.write(error_message)
        print(e)
    

for search_query in keywords:
    try:
        main(search_query)
    except Exception as e:
        print(e)
        continue
