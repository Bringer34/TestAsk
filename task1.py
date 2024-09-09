from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from fake_useragent import UserAgent
from selenium.webdriver.common.action_chains import ActionChains
import time
import csv


useragent = UserAgent()
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument(f"user-agent={useragent.random}")
options.add_argument("--disable-blink-features")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
driver = webdriver.Chrome(options=options)
url = 'https://www.nseindia.com/'
try:
    driver.get(url)
    hover = driver.find_element(By.ID, 'link_2')
    ActionChains(driver).move_to_element(hover).perform()
    time.sleep(1) #небольшое ожидание после наведение на hover
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//a[text()='Pre-Open Market']"))).click() #Клик по Pre-Open Market
    time.sleep(3) # ожидание прогрузки страницы
    q = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//a[@class='symbol-word-break']"))) #ожидаем наличие элемента и парсим
    w = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//table[@id='livePreTable']/tbody/tr/td[7]")))

    #Парсим элементы symbol и final
    Symbol = driver.find_elements(By.XPATH, "//table[@id='livePreTable']/tbody/tr/td[2]") # Symbol
    Final = driver.find_elements(By.XPATH, "//table[@id='livePreTable']/tbody/tr/td[7]") # Final
    Symbol.remove(Symbol[-1]) #удаляем пустой элемент из строки Total
    Final.remove(Final[-1]) #удаляем пустой элемент из строки Total
    ArrayForSort = []

    #формируем массив данных, получаем текст значений и сортируем по алфавиту
    for i in range(0, len(Symbol)):
        Symbol[i] = Symbol[i].text
        Final[i] = Final[i].text
        ArrayForSort.append([Symbol[i],Final[i]]) #создаем список с именем и ценой
    ArrayForSort.sort() #сортируем список

    #записываем в csv файл
    with open('File.csv', mode='w', newline='', encoding='Windows-1251') as file:
        csv.writer = csv.writer(file, delimiter=';')
        csv.writer.writerow(["Имя", "Цена"])
        for i in range(0, len(ArrayForSort)):
            csv.writer.writerow(ArrayForSort[i])
    driver.find_element(By.XPATH, "//a[@class='navbar-brand me-auto']").click() # переход на главную страницу
    time.sleep(3) #ожидаение прогрузки
    element = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[11]/div[1]/section[3]/div/div/div[1]/section/div/div/div[1]/figure/a")))
    actions = ActionChains(driver)
    actions.move_to_element(element).perform()
    time.sleep(0.5)
    driver.find_element(By.XPATH, "//div[@id='CA_Slider']/ul/li/button[@aria-label='Slide 2']").click() #клик по 2 элементу
    time.sleep(1.5)
    driver.find_element(By.XPATH, "//div[@id='CA_Slider']/ul/li/button[@aria-label='Slide 3']").click() #клик по 3 элементу
    time.sleep(1.5)
    driver.execute_script("window.scrollTo(0, 0);") #возвращаемся наверх
    time.sleep(2)
    driver.find_element(By.XPATH, "//a[@id='tabList_NIFTYBANK']").click()
    time.sleep(1.5)
    driver.find_element(By.ID, "view_all_indices").click()
    time.sleep(5)
    element = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//a[text()='NIFTY FINANCIAL SERVICES 25/50']")))#поиск элемента Services 25/50
    actions = ActionChains(driver)
    actions.move_to_element(element).perform() #прокрутка до элемента Services 25/50
    time.sleep(3)
    driver.find_element(By.XPATH, "//a[text()='NIFTY BANK']").click() #кликаем по NIFTY BANK
    time.sleep(5)
    WebDriverWait(driver, 5).until(EC.presence_of_element_located(
        (By.XPATH, "//a[text()='NIFTY FINANCIAL SERVICES 25/50']"))).click()  # выбор элемента Services 25/50 из выпадающего списка
    time.sleep(4) # ожидание перед завершением программы
except Exception as ex:
    print(ex)
#выключаем драйвер
driver.close()
driver.quit()

