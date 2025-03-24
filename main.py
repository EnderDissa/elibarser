from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
import time
from bs4 import BeautifulSoup
import requests
import pandas as pd
import random

def extract_names_from_table(file_path):
    df = pd.read_excel(file_path, header=3, usecols=[0])
    # full_names = []
    publications = []

    for index, row in df.iterrows():
        # surname = row[0]
        # first_name = row[1]
        # patronymic = row[2]
        # initials = f"{first_name[0]}.{patronymic[0]}."
        # full_name = f"{surname} {initials}"
        # full_names.append((surname, first_name, patronymic))
        pubications_name = row[0]
        publications.append(pubications_name)
    return publications

options = Options()
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.5',
    'Connection': 'keep-alive',
    'Referer': 'https://elibrary.ru/'
}
options.set_preference("general.useragent.override", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
options.headless = False

def calc_sleep(SLEEP=1):
    SLEEPY = SLEEP * random.uniform(1, 3)
    print(f"Sleeping for {SLEEPY} seconds.")
    return SLEEPY

def run_script( pub, begin_year="2023", end_year="2025", reset_checkboxes=True):
    url = 'https://elibrary.ru/querybox.asp'
    driver.get(url)

    if "page_captcha.asp" in driver.current_url:
        print("Captcha detected, waiting for user to solve...")
        time.sleep(15)
        driver.refresh()
        return "CAPTCHA"

    time.sleep(calc_sleep())

    # authors = driver.find_elements(By.XPATH, '//select[@name="authors"]/option')
    # if authors:
    #     author = authors[0]
    #     author.click()
    #
    # delete_buttons = driver.find_elements(By.XPATH, '//a[text()="Удалить"]')
    # if len(delete_buttons) > 1:
    #     delete_button = delete_buttons[1]
    #     delete_button.click()
    # time.sleep(calc_sleep())
    #
    # add_buttons = driver.find_elements(By.XPATH, '//a[text()="Добавить"]')
    # if len(add_buttons) > 1:
    #     add_button = add_buttons[1]
    #     add_button.click()
    search_input = driver.find_element(By.NAME, 'ftext')  # Здесь мы меняем поиск по автору на поиск по статье
    search_input.clear()
    search_input.send_keys(pub)

    time.sleep(calc_sleep())

    # windows = driver.window_handles
    # driver.switch_to.window(windows[1])
    # search_input = driver.find_element(By.NAME, 'qwd')
    # search_input.send_keys(author_name)
    # search_link = driver.find_element(By.XPATH, '//a[text()="Поиск"]')
    # search_link.click()
    # time.sleep(calc_sleep())

    #
    # authors = driver.find_elements(By.XPATH, '//td[@align="left" and @valign="middle"]/a')
    #
    # for i in range(len(authors)):
    #     works_count = driver.find_elements(By.XPATH,
    #                                        f'//a[text()="{authors[i].text.strip()}"]/../../following-sibling::tr/td/font[@color="#00008f"]')
    #     if works_count:
    #         num_works = int(works_count[0].text.strip())
    #         if num_works > 15:
    #             print(f"У {authors[i].text.strip()} больше 15 работ ({num_works}). Переход к следующему автору.")
    #             if i == len(authors) - 1:
    #                 return
    #             continue
    #     authors[i].click()
    #     time.sleep(calc_sleep())
    #     break
    #
    # driver.switch_to.window(windows[0])

    begin_year_select = driver.find_element(By.NAME, "begin_year")
    end_year_select = driver.find_element(By.NAME, "end_year")

    begin_year_select.send_keys(begin_year)
    end_year_select.send_keys(end_year)

    article_checkbox = driver.find_element(By.NAME, 'type_article')
    title_checkbox = driver.find_element(By.NAME, 'where_name')
    # book_checkbox = driver.find_element(By.NAME, 'type_book')
    conf_checkbox = driver.find_element(By.NAME, 'type_conf')

    if not article_checkbox.is_selected():
        driver.find_element(By.NAME, 'type_article').click()
    if not title_checkbox.is_selected():
        driver.find_element(By.NAME, 'where_name').click()
        # if not book_checkbox.is_selected():
        #     driver.find_element(By.NAME, 'type_book').click()
    if not conf_checkbox.is_selected():
        driver.find_element(By.NAME, 'type_conf').click()

    search_button = driver.find_element(By.XPATH, '//a[text()="Поиск"]')
    search_button.click()
    time.sleep(calc_sleep())

    html_content = driver.page_source
    soup = BeautifulSoup(html_content, 'html.parser')

    book_links = soup.find_all('a', href=True)
    book_urls = {}

    for link in book_links:
        href = link['href']
        if '/item.asp?id=' in href:
            full_url = 'https://www.elibrary.ru' + href
            title_tag = link.find('span', style="line-height:1.0;")
            title = title_tag.text.strip() if title_tag else "Без названия"

            author_tag = link.find_previous('font', {'color': '#00008f'})
            author_name = "Не найдено"
            if author_tag:
                author_text = author_tag.text.strip()
                author_name = author_text.split(',')[0].strip() if author_text else "Не найдено"

        book_urls[full_url] = {'title': title, 'author': author_name, 'url': full_url}

    return book_urls


driver = webdriver.Firefox(options=options)

publications = extract_names_from_table("вход.xlsx")
print(publications)
for pub in publications:
    try:
        results = []
        result = run_script( pub, begin_year="2023", end_year="2025")
        first_time_flag = 0
        if result == "CAPTCHA":
            print(f"CAPTCHA обнаружена для {pub}, ждем, пока она будет решена...")
            while result == "CAPTCHA":
                result = run_script( pub, begin_year="2023", end_year="2025")

        if result:
            for book_url, title, author_name in result.items():
                results.append((title, author_name, book_url))
            df_results = pd.DataFrame(results, columns=["Название статьи","Первый автор", "Ссылка на статью"])

            try:
                existing_df = pd.read_excel("вывод.xlsx")
                with pd.ExcelWriter("вывод.xlsx", engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    df_results.to_excel(writer, index=False, header=False, startrow=len(existing_df) + 1)
            except FileNotFoundError:
                df_results.to_excel("вывод.xlsx", index=False)
        else:
            print(f"Для {pub} не найдено подходящих работ.")
    except Exception as e:
        print("страшная вонючая ошибка " + str(e))
print("Результаты сохранены в файл 'вывод.xlsx'.")
