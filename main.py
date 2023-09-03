import glob
import os
import win32com.client
import Levenshtein as lev
import re
from selenium import webdriver
import time

# test area
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

chrome_options = Options()
chrome_options.add_experimental_option("detach", True)

def parseData():
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
            raise Exception(word + 'nie jest na liście sportów. Zmień harmonogram.')

    word = win32com.client.Dispatch("Word.Application")
    word.visible = 0

    for i, doc in enumerate(glob.iglob("*.doc")):
        in_file = os.path.abspath(doc)
        wb = word.Documents.Open(in_file)
        out_file = os.path.abspath("out.docx".format(i))
        wb.SaveAs2(out_file, FileFormat=16) # file format for docx
        wb.Close()

    word.Quit()

    from docx import Document

    # Load the first table from your document. In your example file,
    # there is only one table, so I just grab the first one.
    document = Document('out.docx')

    # Data will be a list of rows represented as dictionaries
    # containing each row's data.
    data = []

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
                data.append([row_data['Data'][:10], row_data['Liczba godzin'], row_data['Godziny zajęć'], spellcheck(row_data['Tematyka zajęć'])])

    os.remove('out.docx')
    if(sum([int(i[1]) for i in data]) != 50):
        raise Exception('Suma godzin nie wynosi 50. Zmień harmonogram.')
    else:
        return data

print(parseData())