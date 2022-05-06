import os
import sys
import re
import time
import pandas as pd
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains


def get_scrolled_driver_selenium(webdriver_path, url, pause_time_sec, DEBUG):
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.use_chromium = True
    driver = webdriver.Chrome()
    driver.implicitly_wait(30)
    driver.get(url)

    WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.CLASS_NAME, 'ListCassette__images__innerFit')))
    height = driver.execute_script("return document.body.scrollHeight")
    height_pre = 0

    while height_pre != height:
        height_pre = height
        elems_article = driver.find_elements_by_class_name('ListBukken__list__item')
        last_elem = elems_article[-1]
        actions = ActionChains(driver)
        actions.move_to_element(last_elem)
        actions.perform()

        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ListCassette__images__innerFit')))
        time.sleep(pause_time_sec)

        height = driver.execute_script("return document.body.scrollHeight")

        if DEBUG:
            break

    print("ページの読み込みが終了しました。")

    return driver


def create_csv_from_whole_property(keyword, driver, csv_dir_path):
    df = pd.DataFrame(columns=[
        'ID',
        '物件名',
        '交通',
        '所在地',
        '築年数',
        '総戸数',
        '建物階',
        'チェック',
        '価格（万円）',
        '管理費（円）',
        '修繕積立金（円）',
        '間取り',
        '専有面積（m2）',
        'URL'
    ])

    html = driver.page_source.encode('utf-8')
    soup = BeautifulSoup(html, 'html.parser')
    results = soup.find_all("li", class_="ListBukken__list__item")
    print("物件数：{}".format(len(results)))
    index = 0

    for result in results:
        index += 1  
        property_name = result.find('h2', class_='ListCassette__title__text') 

        price = result.find('dd', class_='ListCassette__information__detail--price')  
        price = (price.text).split("、")[0][:-2] 
        if "億" in price: 
            price = price.split("億")
            price = int(price[0]) * 10000 + int(price[1].replace(',', ''))
        else:
            price = int(price.replace(',', ''))

        check = result.find('h2', class_='ListCassette__title__text--inner')  
        url_ = [url.get('href') for url in result.find_all('a')][0]  

        detail_info = result.find_all('dd', class_='ListCassette__information__detail')
        total_units = re.sub(r"\D", "", detail_info[3].text)  

        tmp_list_1 = (detail_info[5].text).split("　")
        management_fee = re.sub(r"\D", "", tmp_list_1[0])  
        maintenance_charge = re.sub(r"\D", "", tmp_list_1[1])  

        tmp_list_2 = (detail_info[6].text).split("/")
        layout = tmp_list_2[0]  
        occupation_area = tmp_list_2[1][:-2] 

        try: 
            df = df.append(
                {
                    'ID': str(index),
                    '物件名': property_name.text,
                    '交通': detail_info[0].text,
                    '所在地': detail_info[1].text,
                    '築年数': detail_info[2].text,
                    '総戸数': total_units,
                    '建物階': detail_info[4].text,
                    'チェック': check.text,
                    '価格（万円）': str(price),
                    '管理費（円）': management_fee,
                    '修繕積立金（円）': maintenance_charge,
                    '間取り': layout,
                    '専有面積（m2）': occupation_area,
                    'URL': url_

                }, ignore_index=True
            )
        
        except Exception as e:
            print("ID:{} のスクレイピングに失敗しました。(URL: {})".format(index, url_))
            print("エラーメッセージ： {}".format(e))

    df.set_index('ID', inplace=True)  

    csv_path = os.path.join(csv_dir_path, keyword + '_物件_一覧ページ.csv')
    df.to_csv(csv_path, encoding='cp932')
    print(csv_path + " が保存されました。")

    df = df[['物件名','URL']]  

    return df


def scraping_and_create_csv_from_each_property(keyword, driver, csv_dir_path, df, pause_time_sec, DEBUG):
    driver.execute_script("window.open()")
    new_window = driver.window_handles[-1]
    driver.switch_to.window(new_window)

    df2 = pd.DataFrame(columns=[
        'ID',
        '物件名',
        '取扱会社',
        '画像枚数',
        '物件内画像枚数順位'
    ])

    df_tmp = df2

    for index_, property_name_, url_ in zip(df.index, df['物件名'], df['URL']): 
        driver.get(url_)
        try:
            WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.CLASS_NAME, 'DetailCompanyInfo2__summaryButtonArea__button__count')))
            time.sleep(pause_time_sec/5)

            html = driver.page_source.encode('utf-8')
            soup = BeautifulSoup(html, 'html.parser')
            results = soup.find_all("div", class_="DetailCompanyInfo2--ag")

            for result in results:
                company = result.find('div', class_='DetailCompanyInfo2__companyName') 
                number_image = result.find('span', class_='DetailCompanyInfo2__summaryButtonArea__button__count')  

                df_tmp = df_tmp.append(
                    {
                        'ID':index_,
                        '物件名':property_name_,
                        '取扱会社': company.text,
                        '画像枚数': int(number_image.text),
                    }, ignore_index=True
                )

            df_tmp = df_tmp.sort_values('画像枚数', ascending=False)  
            df_tmp['物件内画像枚数順位'] = df_tmp['画像枚数'].rank(ascending=False, method='min').astype('int')  

            df2 = pd.concat([df2, df_tmp])  
            df_tmp = df_tmp.drop(df_tmp.index[:])  

            print("ID:{} {} のスクレイピングが終了しました。(URL: {})".format(index_, property_name_, url_))

        except Exception as e:
            print("ID:{} {} のスクレイピングに失敗しました。(URL: {})".format(index_, property_name_, url_))
            print("エラーメッセージ： {}".format(e))

        if DEBUG:  
            if int(index_) >= 3:
                break

    driver.quit() 
    csv_path2 = os.path.join(csv_dir_path, keyword + '_物件_個別ページ.csv')
    df2.to_csv(csv_path2, encoding='cp932', index=False)
    print(csv_path2 + " が保存されました。")

    df3 = df2[df2['取扱会社'].str.contains(keyword)]  
    csv_path3 = os.path.join(csv_dir_path, keyword + '_物件内画像枚数順位.csv')
    df3.to_csv(csv_path3, encoding='cp932', index=False)
    print(csv_path3 + " が保存されました。")


if __name__ == "__main__":
    os.chdir(os.path.dirname(os.path.abspath(__file__)))  
    webdriver_path = ".\\chromedriver_win32\\chromedriver.exe"  
    csv_dir_path = ".\\csv"  
    os.makedirs(csv_dir_path, exist_ok=True) 
    keyword = "エステート白馬 東大宮" 
    url = "https://realestate.yahoo.co.jp/new/house/search/03/?query=" + keyword 
    pause_time_sec = 5  

    DEBUG = False

    driver = get_scrolled_driver_selenium(webdriver_path, url, pause_time_sec, DEBUG)
    df = create_csv_from_whole_property(keyword, driver, csv_dir_path)
    scraping_and_create_csv_from_each_property(keyword, driver, csv_dir_path, df, pause_time_sec, DEBUG)
