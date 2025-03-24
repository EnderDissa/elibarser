from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
import time
from bs4 import BeautifulSoup
import requests
import pandas as pd
import random

def extract_names_from_table(file_path):
    df = pd.read_excel(file_path, header=3, usecols=[0, 1, 2])
    full_names = []

    for index, row in df.iterrows():
        surname = row[0]
        first_name = row[1]
        patronymic = row[2]
        initials = f"{first_name[0]}.{patronymic[0]}."
        full_name = f"{surname} {initials}"
        full_names.append((surname, first_name, patronymic))

    return full_names

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
    print("Sleeping for {} seconds.".format(SLEEPY))
    return SLEEPY

def run_script(first_time_flag, author_name, begin_year="2023", end_year="2025", reset_checkboxes=True):
    url = 'https://elibrary.ru/querybox.asp'
    driver.get(url)

    if "page_captcha.asp" in driver.current_url:
        print("Captcha detected, waiting for user to solve...")
        time.sleep(15)
        driver.refresh()
        return "CAPTCHA"

    time.sleep(calc_sleep())

    authors = driver.find_elements(By.XPATH, '//select[@name="authors"]/option')
    if authors:
        author = authors[0]
        author.click()

    delete_buttons = driver.find_elements(By.XPATH, '//a[text()="Удалить"]')
    if len(delete_buttons) > 1:
        delete_button = delete_buttons[1]
        delete_button.click()
    time.sleep(calc_sleep())

    add_buttons = driver.find_elements(By.XPATH, '//a[text()="Добавить"]')
    if len(add_buttons) > 1:
        add_button = add_buttons[1]
        add_button.click()

    time.sleep(calc_sleep())
    windows = driver.window_handles
    driver.switch_to.window(windows[1])
    search_input = driver.find_element(By.NAME, 'qwd')
    search_input.send_keys(author_name)
    search_link = driver.find_element(By.XPATH, '//a[text()="Поиск"]')
    search_link.click()
    time.sleep(calc_sleep())


    authors = driver.find_elements(By.XPATH, '//td[@align="left" and @valign="middle"]/a')

    for i in range(len(authors)):
        works_count = driver.find_elements(By.XPATH,
                                           f'//a[text()="{authors[i].text.strip()}"]/../../following-sibling::tr/td/font[@color="#00008f"]')
        if works_count:
            num_works = int(works_count[0].text.strip())
            if num_works > 15:
                print(f"У {authors[i].text.strip()} больше 15 работ ({num_works}). Переход к следующему автору.")
                if i == len(authors) - 1:
                    return
                continue
        authors[i].click()
        time.sleep(calc_sleep())
        break

    driver.switch_to.window(windows[0])

    begin_year_select = driver.find_element(By.NAME, "begin_year")
    end_year_select = driver.find_element(By.NAME, "end_year")

    begin_year_select.send_keys(begin_year)
    end_year_select.send_keys(end_year)

    if reset_checkboxes:
        article_checkbox = driver.find_element(By.NAME, 'type_article')
        book_checkbox = driver.find_element(By.NAME, 'type_book')
        conf_checkbox = driver.find_element(By.NAME, 'type_conf')

        if not article_checkbox.is_selected():
            driver.find_element(By.NAME, 'type_article').click()
        if not book_checkbox.is_selected():
            driver.find_element(By.NAME, 'type_book').click()
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
            book_urls[full_url] = title
    time.sleep(calc_sleep())

    return book_urls

first_time_flag = 1
driver = webdriver.Firefox(options=options)

names = extract_names_from_table("вход.xlsx")
print(names)
for name in names:
    try:
        results = []
        author_name = f"{name[0]} {name[1][0]}.{name[2][0]}."
        print(author_name)
        result = run_script(first_time_flag, author_name, begin_year="2023", end_year="2025")
        first_time_flag = 0
        if result == "CAPTCHA":
            print(f"CAPTCHA обнаружена для {author_name}, ждем, пока она будет решена...")
            while result == "CAPTCHA":
                result = run_script(first_time_flag, author_name, begin_year="2023", end_year="2025")

        if result:
            for book_url, title in result.items():
                results.append((name[0], name[1], name[2], title, book_url))
            df_results = pd.DataFrame(results, columns=["Фамилия", "Имя", "Отчество", "Название статьи", "Ссылка на статью"])

            try:
                existing_df = pd.read_excel("вывод.xlsx")
                with pd.ExcelWriter("вывод.xlsx", engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    df_results.to_excel(writer, index=False, header=False, startrow=len(existing_df) + 1)
            except FileNotFoundError:
                df_results.to_excel("вывод.xlsx", index=False)
        else:
            print(f"Для {author_name} не найдено подходящих работ или есть больше 15 работ.")
    except Exception as e:
        print("страшная вонючая ошибка " + str(e))


print("Результаты сохранены в файл 'вывод.xlsx'.")
