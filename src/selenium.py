from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.firefox.service import Service as FirefoxService
import time
from datetime import datetime
import re
import random
import ctypes

def genOptions(sportType):
    options = { # szablon formularza
    'rodzaj' : 'Trening',
    'mozna_dolaczyc' : 'Tak',
    'niepelnosprawni' : 'Nie',
    'obcokrajowcy' : 'Nie',
    'miejsce' : 'Boisko piłkarskie',
    'grupy_wiekowe' : ['Szkoła Podstawowa', 'Szkoła Ponadpodstawowa', 'dorośli'],
    'plec' : ['dziewczęta / kobiety', 'chłopcy / mężczyźni'],
    'dyscyplina' : 'Piłka nożna',
    }

    if sportType == 'Piłka nożna':
        options['grupy_wiekowe'] = ['Szkoła Podstawowa', 'Szkoła Ponadpodstawowa']
            
    if sportType == 'Wiele dyscyplin':
        options['rodzaj'] = 'Animacja / gry i zabawy'
        options['miejsce'] = 'Cały obiekt'

    if sportType in ['Koszykówka', 'Tenis']:
        options['miejsce'] = 'Boisko wielofunkcyjne'

    if sportType in ['Gra w bule', 'Bule']:
        options['rodzaj'] = 'Animacja / gry i zabawy'
        options['miejsce'] = 'Cały obiekt'
        options['grupy_wiekowe'] = ['Szkoła Podstawowa', 'Szkoła Ponadpodstawowa', 'dorośli', 'seniorzy (60+)']

    if sportType != 'Wiele dyscyplin':
        options['dyscyplina'] = sportType
    
    if random.randint(0, 10) == 10: # xD
        options['niepelnosprawni'] = 'Tak'
    return options

def sumSplit(left,right=[],difference=0):
    sumLeft,sumRight = sum(left),sum(right)

    # stop recursion if left is smaller than right
    if sumLeft<sumRight or len(left)<len(right): return

    # return a solution if sums match the tolerance target
    if sumLeft-sumRight == difference:
        return left, right, difference

    # recurse, brutally attempting to move each item to the right
    for i,value in enumerate(left):
        solution = sumSplit(left[:i]+left[i+1:],right+[value], difference)
        if solution: return solution

    if right or difference > 0: return 
    # allow for imperfect split (i.e. larger difference) ...
    for targetDiff in range(1, sumLeft-min(left)+1):
        solution = sumSplit(left, right, targetDiff)
        if solution: return solution 

miesiace = ['styczeń', 'luty', 'marzec', 'kwiecień', 'maj', 'czerwiec', 'lipiec', 'sierpień', 'wrzesień', 'październik', 'listopad','grudzień']

with open('auth.txt', 'r') as f:
    auth = f.read().split(':')

with open('data.txt', 'r', encoding='utf-8') as f:
    data = f.read().split('\n')
data.pop() # remove last newline

entries = []
liczba_godz = []
for i in data:
    entries.append(i.split(';'))
    liczba_godz.append(int(i.split(';')[1]))

msit = sumSplit(liczba_godz)[0]
jst = sumSplit(liczba_godz)[1]
diff = sumSplit(liczba_godz)[2]

if diff != 0:
    entries.append(entries[-1][0], diff, '10-' + str(diff), random.choice(['Mini turniej'], ['Gry i zabawy'], ['Trening']), 'Wiele dyscyplin')
    if(sum(msit) > sum(jst)):
        msit.append(diff)
    else:
        jst.append(diff)

# Start driver
service = ChromeService('../src/drivers/chromedriver.exe')
driver = webdriver.Chrome(service=service)
driver.get('https://system.programorlik.pl/kalendarz/kalendarz')

time.sleep(1)

# Get login page
email = driver.find_element("name", "email")
password = driver.find_element("name", "password")
email.send_keys(auth[0])
password.send_keys(auth[1])
driver.find_element("class name", "btn-primary").click()

# Get calendar page
time.sleep(1)
cur_month = miesiace.index(re.sub(r'[^a-zA-Ząęłśćóńź]+', r'', driver.find_element("class name", "fc-toolbar-title").text))

data_date = datetime.strptime(data[0][0:10], '%d.%m.%Y')
month_discrepancy = data_date.month - cur_month-1 # różnica między obecnym miesiącem a miesiącem zapisanym w harmonogramie
firstRun = True

# Wybierz odpowiedni miesiąc
for i in entries:
    if month_discrepancy != 0:
        if(month_discrepancy) < 0:
            for j in range(abs(month_discrepancy)):
                driver.find_element("class name", "fc-prev-button").click()
        else:
            for j in range(abs(month_discrepancy)):
                driver.find_element("class name", "fc-next-button").click()
    
    if len(driver.find_elements("class name", "fc-event-title")) > 0 and firstRun:
        ctypes.windll.user32.MessageBoxW(0, "W tym miesiącu są już zapisane zajęcia. Usuń je i spróbuj ponownie.", "Projekt Orlik", 1)
        driver.close()
        exit()

    time.sleep(1)
    # Wypełnij formularz

    # Wybierz odpowiedni dzień
    cur_date = datetime.strptime(i[0], '%d.%m.%Y')
    button_name = cur_date.strftime('%Y-%m-%d')
    button = driver.find_element("xpath", f'//td[@data-date="{button_name}"]') 
    button.click()
    time.sleep()

    # Wypełnij dane
    # https://example.com/form_layout.jpg

    options = genOptions(i[4]) # Wygeneruj poprawne opcje do wpisania dla formularza zgodnie z nazwą sportu

    # Tytuł
    driver.find_element("id", "title").send_keys(i[3])
    
    # Rodzaj wydarzenia
    if(options['rodzaj'] != 'Trening'):
        driver.find_element("xpath", "//button[@class='btn text-pre-line calendar_selectpicker dropdown-toggle btn-light']").click()
        time.sleep(0.5)
        driver.find_element("xpath", f"//span[@class='text'][contains(.,'{options['rodzaj']}')]").click()

    # Godziny start-stop
    godziny = []
    for j in i[2].split('-'):
        if j.find(':') == -1:
            j += ':00'
        godziny.append(j)
    driver.find_element("id", "godzina_od").send_keys(godziny[0])
    driver.find_element("id", "godzina_do").send_keys(godziny[1])

    # Finansowane przez
    l_godz = int(i[1])
    if l_godz in msit:
        driver.find_element("id", "IDFINANSOWANIE1").click() # MSiT
        msit.remove(l_godz)
    elif l_godz in jst:
        driver.find_element("id", "IDFINANSOWANIE2").click() # JST
        jst.remove(l_godz)

    # Niepełnosprawni
    if(options['niepelnosprawni'] == 'Tak'):
        driver.find_element("id", "NIEPELNOSPRAWNI1").click()
    else:
        driver.find_element("id", "NIEPELNOSPRAWNI0").click()

    # Obcokrajowcy
    driver.find_element("id", "obcokrajowcy0").click()

    # Miejsce zajęć
    if(options['miejsce'] != 'Boisko piłkarskie'):
        rodzaj = driver.find_element("xpath", '//button[@data-id="idlokalizacja"]').click()
        time.sleep(0.2)
        driver.find_element("xpath", f"//span[@class='text'][contains(.,'{options['miejsce']}')]").click()

    # Grupy wiekowe
    for j in options['grupy_wiekowe']:
        driver.find_element("xpath", f"//select[@id='idgrupa_wiekowa']/option[contains(.,'{j}')]").click()

    # Płeć
    for j in options['plec']:
        driver.find_element("xpath", f"//select[@id='idplec']/option[contains(.,'{j}')]").click()

    # Dyscyplina
    driver.find_element("xpath", '//button[@data-id="iddyscypliny"]').click()
    time.sleep(0.2)
    driver.find_element("xpath", f"//span[@class='text'][contains(.,'{options['dyscyplina']}')]").click()

    # Zapisz

    driver.find_element("xpath", "//div[@class='bottom_container']/button[@id='btn_next12']").click()
    time.sleep(0.5)
    driver.find_element("xpath", "//html/body[@class='sidebar-mini modal-open']/div[@class='wrapper']/div[@class='content-wrapper ']/div[@class='content']/div[@class='container-fluid']/div[@id='dialogConfirm']/div[@class='modal-dialog modal-dialog-centered']/div[@class='modal-content']/div[@class='modal-footer']/button[@id='dialogConfirmButton']").click()
    firstRun = False

driver.close()
ctypes.windll.user32.MessageBoxW(0, "Program zakończony pomyślnie.", "Projekt Orlik", 1)