from urllib.request import urlopen, Request
from bs4 import BeautifulSoup
import openpyxl as xl
from openpyxl.styles import Font


webpage = 'https://registrar.web.baylor.edu/exams-grading/spring-2023-final-exam-schedule'
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.3'}

req = Request(url=webpage, headers=headers)

page = urlopen(req)

soup = BeautifulSoup(page, 'html.parser')

title = soup.title

print(title.text)

myclasses = ['MW 1:00 p.m.', 'MW 2:30 p.m.', 'MW 4:00 p.m.', 'TR 11:00 a.m.']

finals_rows = soup.findAll('tr')

for row in finals_rows:
    final = row.findAll('td')
    if final: 
        myclass = final[0].text
        if myclass in myclasses:
            print()
            print(f'For class: {myclass} the final is scheduled for {final[1].text} at {final[2].text}')
            print()