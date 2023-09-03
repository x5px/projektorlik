from selenium import webdriver
import time
import re

miesiace = ['styczeń', 'luty', 'marzec', 'kwiecień', 'maj', 'czerwiec', 'lipiec', 'sierpień', 'wrzesień', 'październik', 'listopad','grudzień']

with open('auth.txt', 'r') as f:
    auth = f.read().split(':')

with open('data.txt', 'r') as f:
    data = f.read().split('\n')
data.pop() # remove last newline

# test area
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service


# Start driver
chrome_options = Options()
chrome_options.add_experimental_option("detach", False) # True to keep window open
service = Service(executable_path=r'../src/drivers/chromedriver.exe')
driver = webdriver.Chrome(service=service, options=chrome_options)
driver.get('https://system.programorlik.pl/kalendarz/kalendarz')

time.sleep(2)

# Get login page
email = driver.find_element("name", "email")
password = driver.find_element("name", "password")
email.send_keys(auth[0])
password.send_keys(auth[1])
driver.find_element("class name", "btn-primary").click()

# Get calendar page
time.sleep(2)
cur_month = miesiace.index(re.sub(r'[^a-zA-Ząęłśćóńź]+', r'', driver.find_element("class name", "fc-toolbar-title").text))

month_discrepancy = int(data[0][4]) - cur_month-1 # różnica między obecnym miesiącem a miesiącem zapisanym w harmonogramie

if month_discrepancy != 0:
    if(month_discrepancy) < 0:
        for i in range(abs(month_discrepancy)):
            driver.find_element("class name", "fc-prev-button").click()
    else:
        for i in range(abs(month_discrepancy)):
            driver.find_element("class name", "fc-next-button").click()
