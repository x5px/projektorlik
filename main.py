import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter import messagebox as mb
from threading import Thread
import os
import win32com.client as win32
from win32com.client import constants
from docx import Document
import Levenshtein as lev
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.firefox.service import Service as FirefoxService
from time import sleep
from datetime import datetime
from random import randint, choice
import ctypes

# GUI
root = tk.Tk()
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("5px.projektorlik") # icon fix
root.iconbitmap('icon.ico')
root.title("Projekt Orlik")
root.geometry("400x200")
root.resizable(False, False)
root.eval('tk::PlaceWindow . center')

data = []

def showLoading(filename):
    
    global l, open_button, l2
    l.pack_forget()
    open_button.pack_forget()
    l2.pack_forget()
    root.update()
    
    l3 = tk.Label(root, text="Ładowanie...")
    l3.config(font=("Calibri", 20))
    l3.grid(row=0, column=0, columnspan=2, pady=10, padx=10)
    pb = ttk.Progressbar(root, orient="horizontal", length=200, mode="determinate")
    pb.grid(row=1, column=0, columnspan=2, pady=10, padx=10)
    l4 = tk.Label(root, text="Proszę czekać. Nie ruszaj myszką podczas działania programu.")
    l4.config(font=("Calibri", 11))
    l4.grid(row=2, column=0, columnspan=2, pady=10, padx=10)

    # run test_function in another thread
    t = Thread(target=parseData, args=(filename,))
    t.start()
    for _ in range(5):
        root.update_idletasks()
        pb.step(20)
        root.after(1000)
    t.join()
    root.withdraw()
    fillPage(data)
    
    
def select_file():
    filetypes = (
        ('Dokument Word 2007', '*.doc'),
        ('Dokument Word', '*.docx'),
        ('Wszytkie pliki', '*.*')
    )

    filename = fd.askopenfilename(
        title='Wybierz harmonogram',
        initialdir=r'%userprofile%/Desktop/harmonogramy',
        filetypes=filetypes
        )

    if(len(filename) > 0):
        showLoading(os.path.abspath(filename))

# Parser

def save_as_docx(path):
    # Opening MS Word
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(path)
    doc.Activate()

    # Rename path with .docx

    # Save and Close
    new_file_abs = os.getcwd() + '/out.docx'

    # Save and Close
    word.ActiveDocument.SaveAs(
        new_file_abs, FileFormat=constants.wdFormatXMLDocument
    )
    doc.Close(False)
    word.Quit()

# Selenium

def parseData(filename):
    global data
    sporty = []
    with open('src/sporty.txt', 'r', encoding='utf-8') as f:
        for line in f:
            sporty.append(line[:-1])

    pattern = r'^([\w]+ [\w]+)' # split second space
    def split_second_space(word):
        if(len(re.findall(pattern, word)) == 0):
            return word
        else:
            return re.match(pattern, word).group()
        
    def spellcheck(word):
        if split_second_space(word).capitalize() in sporty:
            return split_second_space(word).capitalize()
        else:
            for i in sporty:
                if lev.distance(i, split_second_space(word).capitalize()) < 3:
                    return i
            os.remove('out.docx')
            raise ValueError(word + 'nie jest na liście sportów. Zmień harmonogram.')

    try:
        print(filename)
        save_as_docx(filename)
    except OSError or FileNotFoundError:
        raise FileNotFoundError('Nie znaleziono pliku harmonogramu. Upewnij się, że plik znajduje się w tym samym folderze co plik main.py.')

    # Load the first table from your document. In your example file,
    # there is only one table, so I just grab the first one.
    document = Document('out.docx')

    keys = None
    for table in document.tables:
        for i, row in enumerate(table.rows):
            text = (cell.text for cell in row.cells)

            # Establish the mapping based on the first row
            # headers; these will become the keys of our dictionary
            if i == 0:
                keys = tuple(text)
                continue

            # Construct a dictionary for this row, mapping
            # keys to values for this row
            row_data = dict(zip(keys, text))
            if(row_data['Godziny zajęć'] not in ['-', '']):
                data.append([row_data['Data'][:10], row_data['Liczba godzin'], row_data['Godziny zajęć'], row_data['Tematyka zajęć'], spellcheck(row_data['Tematyka zajęć'])])

    os.remove('out.docx')
    if(sum([int(i[1]) for i in data]) < 50):
        # print(sum([int(i[1]) for i in data]))
        raise AssertionError('Suma godzin jest mniejsza niż 50. Zmień harmonogram.')
    else:
        return

# Selenium
def fillPage(data):
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
        
        if randint(0, 10) == 10: # xD
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

    liczba_godz = []
    for i in data:
        liczba_godz.append(int(i[1]))

    msit = sumSplit(liczba_godz)[0]
    jst = sumSplit(liczba_godz)[1]
    diff = sumSplit(liczba_godz)[2]

    if diff != 0:
        data.append(data[-1][0], diff, '10-' + str(diff), choice(['Mini turniej'], ['Gry i zabawy'], ['Trening']), 'Wiele dyscyplin')
        if(sum(msit) > sum(jst)):
            msit.append(diff)
        else:
            jst.append(diff)

    # Start driver
    service = ChromeService('drivers/chromedriver.exe')
    driver = webdriver.Chrome(service=service)
    driver.get('https://system.programorlik.pl/kalendarz/kalendarz')

    sleep(1)

    # Get login page
    email = driver.find_element("name", "email")
    password = driver.find_element("name", "password")
    email.send_keys(auth[0])
    password.send_keys(auth[1])
    driver.find_element("class name", "btn-primary").click()

    # Get calendar page
    sleep(2)
    cur_month = miesiace.index(re.sub(r'[^a-zA-Ząęłśćóńź]+', r'', driver.find_element("class name", "fc-toolbar-title").text))

    data_date = datetime.strptime(data[0][0][0:10], '%d.%m.%Y')
    month_discrepancy = data_date.month - cur_month-1 # różnica między obecnym miesiącem a miesiącem zapisanym w harmonogramie
    firstRun = True

    # Wybierz odpowiedni miesiąc
    for i in data:
        sleep(1)
        if month_discrepancy != 0:
            if(month_discrepancy) < 0:
                for j in range(abs(month_discrepancy)):
                    driver.find_element("class name", "fc-prev-button").click()
            else:
                for j in range(abs(month_discrepancy)):
                    driver.find_element("class name", "fc-next-button").click()
        
        # if len(driver.find_elements("class name", "fc-event-title")) > 0 and firstRun:
        #     root.iconify()
        #     # msgbox Tk
        #     mb.showerror('Projekt Orlik', 'W tym miesiącu są już zapisane zajęcia. Usuń je i spróbuj ponownie.')
                    
        #     driver.quit()
        #     root.destroy()
        #     exit()

        # Wypełnij formularz

        # Wybierz odpowiedni dzień
        cur_date = datetime.strptime(i[0], '%d.%m.%Y')
        button_name = cur_date.strftime('%Y-%m-%d')
        button = driver.find_element("xpath", f'//td[@data-date="{button_name}"]') 
        button.click()
        sleep(0.5)

        # Wypełnij dane
        # https://example.com/form_layout.jpg

        options = genOptions(i[4]) # Wygeneruj poprawne opcje do wpisania dla formularza zgodnie z nazwą sportu

        # Tytuł
        driver.find_element("id", "title").send_keys(i[3])
        
        # Rodzaj wydarzenia
        if(options['rodzaj'] != 'Trening'):
            driver.find_element("xpath", "//button[@class='btn text-pre-line calendar_selectpicker dropdown-toggle btn-light']").click()
            sleep(0.2)
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
            driver.find_element("xpath", '//button[@data-id="idlokalizacja"]').click()
            sleep(0.2)
            driver.find_element("xpath", f"//span[@class='text'][contains(.,'{options['miejsce']}')]").click()

        # Grupy wiekowe
        for j in options['grupy_wiekowe']:
            driver.find_element("xpath", f"//select[@id='idgrupa_wiekowa']/option[contains(.,'{j}')]").click()

        # Płeć
        for j in options['plec']:
            driver.find_element("xpath", f"//select[@id='idplec']/option[contains(.,'{j}')]").click()

        # Dyscyplina
        driver.find_element("xpath", '//button[@data-id="iddyscypliny"]').click()
        sleep(0.2)
        driver.find_element("xpath", f"//span[@class='text'][contains(.,'{options['dyscyplina']}')]").click()

        # Zapisz

        driver.find_element("xpath", "//div[@class='bottom_container']/button[@id='btn_next12']").click()
        sleep(0.5)
        driver.find_element("xpath", "//html/body[@class='sidebar-mini modal-open']/div[@class='wrapper']/div[@class='content-wrapper ']/div[@class='content']/div[@class='container-fluid']/div[@id='dialogConfirm']/div[@class='modal-dialog modal-dialog-centered']/div[@class='modal-content']/div[@class='modal-footer']/button[@id='dialogConfirmButton']").click()
        firstRun = False

    driver.quit()
    root.iconify()
    mb.showinfo('Zakończono', 'Program zakończył działanie pomyślnie. Możesz teraz wygenerować rozliczenie.')
    root.destroy()
    exit()

l = tk.Label(root, text="Wybierz harmonogram")
l.config(font=("Calibri", 20))

l2 = tk.Label(root, text="Made by Jakub Rutkowski (5px) 2023")
l2.config(font=("Calibri", 11))
open_button = tk.Button(
    root,
    width=20,
    height=2,
    text='Otwórz plik',
    command=select_file
)

l.pack()
open_button.pack(expand=True)
l2.pack()

# run the application
root.mainloop()
    