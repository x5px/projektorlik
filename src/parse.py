from bs4 import BeautifulSoup

with open('src/picker.html', encoding='utf-8') as html_file:
    html = html_file.read()

soup = BeautifulSoup(html, 'html.parser')
spans = soup.find_all('span', class_='text')

with open('src/sporty.txt', 'a+', encoding='utf-8') as f:
    for span in spans:
        f.write(span.text[:-1] + '\n')