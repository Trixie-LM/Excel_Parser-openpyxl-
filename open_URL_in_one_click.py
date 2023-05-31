import os
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait

# Тестовые данные 1С
# Если надо открывать ссылки на другом стенде, то нужно сменить
# LOGIN = 'Exchange'
# PASSWORD = 'LLAkDL'
#
# driver = webdriver.Firefox()
# driver.get("http://10.240.240.99/Lotteries_Trade11_Piganov/hs/Tickets/Status/Ticket/20630130007681301551")
# wait = WebDriverWait(driver, 30)
# driver.maximize_window()

# driver.find_element(By.XPATH, "//button[ contains(text(), 'ОК') ]").click()
# password_field.send_keys(PASSWORD)
# password_field.send_keys(Keys.RETURN)
# driver.implicitly_wait(10)
# wait.until( EC.element_to_be_clickable((By.XPATH, "//span[ contains(text(), 'Товары и услуги')]")))
# driver.get("https://partners.int.multibonus.sh/catalog/onlinecategory/0e80b6dc-0d48-498d-8671-6064343b4ddc?sort=POPULARITY&direction=NEXT")
# driver.switch_to.frame(driver.find_element(By.XPATH, "//html/body/div/main//iframe"))

def url_in_one_click(fileName):
    try:
        fileName = fileName + '.txt'
        file = open(fileName)
    except FileNotFoundError:
        fileName = 'Расхождения в отчетах/' + fileName
        file = open(fileName)



    for site in file:
        os.system('start ' + site)

    file.close()


