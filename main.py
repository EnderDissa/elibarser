from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
import time
from bs4 import BeautifulSoup
import requests
import pandas as pd

def extract_names_from_table(file_path):
    # Чтение HTML в DataFrame
    df = pd.read_html(file_path, header=0)[0]

    # Если столбец с фамилией и именем называется, например, 'ФИО', замените его название
    # Для примера, допустим, колонка с фамилией и именем называется 'ФИО'
    for index, row in df.iterrows():
        full_name = row['ФИО']  # Пример: замените 'ФИО' на ваш столбец
        name_parts = full_name.split()  # Разделяем ФИО по пробелу
        if len(name_parts) >= 2:
            surname = name_parts[0]  # Фамилия
            initials = " ".join(name_parts[1:3])  # Имя и отчество
            print(f"Фамилия: {surname}, И.О.: {initials}")


options = Options()
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.5',
    'Connection': 'keep-alive',
}
options.headless = False
SLEEP = 5

def run_script(first_time_flag, author_name, reset_checkboxes=True):
    url = 'https://elibrary.ru/querybox.asp'
    driver.get(url)

    if "page_captcha.asp" in driver.current_url:
        print("Captcha detected, waiting for user to solve...")
        time.sleep(10)
        driver.refresh()
        return "CAPTCHA"

    time.sleep(SLEEP)

    if reset_checkboxes:
        # Сбрасываем форму и чекбоксы перед каждым запросом
        driver.find_element(By.NAME, 'surname').clear()  # Очистим поле автор
        driver.find_element(By.NAME, 'surname').send_keys(author_name)  # Вводим нового автора

        # Сбрасываем все чекбоксы
        driver.find_element(By.NAME, 'type_article').click()
        driver.find_element(By.NAME, 'type_book').click()
        driver.find_element(By.NAME, 'type_conf').click()

    add_buttons = driver.find_elements(By.XPATH, '//a[text()="Добавить"]')
    if len(add_buttons) > 1:
        add_button = add_buttons[1]
        add_button.click()
    else:
        print("Не найдено второго элемента 'Добавить'.")
        return

    time.sleep(SLEEP)

    windows = driver.window_handles
    driver.switch_to.window(windows[1])

    search_input = driver.find_element(By.NAME, 'qwd')
    search_input.send_keys(author_name)

    search_link = driver.find_element(By.XPATH, '//a[text()="Поиск"]')
    search_link.click()

    author_link = driver.find_element(By.XPATH, '/html/body/center/form/table[2]/tbody/tr[2]/td[2]/a')
    author_link.click()

    time.sleep(SLEEP)

    driver.execute_script("location.reload(true)")  # Перезагружаем страницу, очищая кэш и сбрасывая чекбоксы

    driver.switch_to.window(windows[0])
    search_button = driver.find_element(By.XPATH, '//a[text()="Поиск"]')
    search_button.click()

    time.sleep(SLEEP)

    html_content = driver.page_source

    soup = BeautifulSoup(html_content, 'html.parser')

    book_links = soup.find_all('a', href=True)
    book_urls = []

    for link in book_links:
        href = link['href']
        if '/item.asp?id=' in href:
            full_url = 'https://www.elibrary.ru' + href
            book_urls.append(full_url)

    print("Ссылки на книги:")
    for url in book_urls:
        print(url)

    doi_list = []

    for book_url in book_urls:
        response = requests.get(book_url, headers=headers)
        page_soup = BeautifulSoup(response.text, 'html.parser')

        doi_tags = page_soup.find_all('a', href=True)
        for doi_tag in doi_tags:
            href = doi_tag['href']
            if "doi.org" in href:
                doi = href.split("doi.org/")[-1]
                doi_list.append(doi)

        time.sleep(SLEEP)

    print("\nDOI для каждой публикации:")
    for doi in doi_list:
        print(doi)

    with open('elibrary_querybox_results.html', 'w', encoding='utf-8') as file:
        file.write(html_content)

    print("HTML страница с результатами успешно сохранена.")

first_time_flag = 1
driver = webdriver.Firefox(options=options)

while True:
    author_name = "Конокин Д В"  # Пример: замени на имя автора, которое нужно искать
    result = run_script(first_time_flag, author_name)
    first_time_flag = 0
    if result == "CAPTCHA":
        first_time_flag = 1
        driver.close()
        driver = webdriver.Firefox(options=options)
        time.sleep(10)
