from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from openpyxl import Workbook

#Robot otwiera stronę w przeglądarce Chrome
s = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=s)
driver.get("https://rpa.hybrydoweit.pl/")

#Robot klika przycisk Artykuły żeby przejść do tej sekcji
button = driver.find_element(By.CSS_SELECTOR, 'ul.rpa-navbar__menu :nth-child(2)')
button.click()

#Robot pobiera tytuły, informacje o dziale/branży i linki
articles = driver.find_elements(By.CSS_SELECTOR, '.rpajs-articles .rpa-article-card .rpa-article-card__title')
articles_meta = driver.find_elements(By.CSS_SELECTOR, '.rpajs-articles .rpa-article-card .rpa-article-card__metadata')
articles_link = driver.find_elements(By.CSS_SELECTOR, '.rpajs-articles .rpa-article-card .rpa-article-card__link')

#Robot pobiera dane o pierwszych i ostatnich pięciu artukułach do listy list
len = len(articles)
list_of_articles = []
def get_article_data(x):
    list_of_articles.append([articles[x].text, articles_meta[x].text.partition(": ")[2], articles_link[i].get_attribute("href")])

for i in range(5):
    get_article_data(i)
for j in range(len-5, len):
    get_article_data(j)

#Zamykamy przeglądarkę
driver.close()

#Robot zapisuje w Excelu pozyskane dane
wb = Workbook()
ws = wb.active
ws.append(["Tytuł", "Branża/Dział", "Link"])

for row in list_of_articles:
    ws.append(row)
wb.save("Spis_arytułów.xlsx")






