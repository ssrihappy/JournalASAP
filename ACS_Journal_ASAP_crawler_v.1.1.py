from selenium import webdriver
import pandas as pd
from bs4 import BeautifulSoup
import requests
from openpyxl import workbook
from io import BytesIO
import requests
import xlsxwriter

# 2021-10-07, v1.1
# update TOC in excel
#


### To-be : merge to one page ###

num = int(input("Please enter the number of journals : ", ))
workbook = xlsxwriter.Workbook('summary.xlsx')
ws = workbook.add_worksheet('TOC')

# TOC setting
ws.set_column('A:B', 4)
ws.set_column('C:C', 32)
ws.set_column('D:D', 100)
ws.set_column('E:E', 10)
ws.set_default_row(80)
ws.set_row(0, 25)
ws.write(0, 0, 'Num')
ws.write(0, 1, 'Abb')
ws.write(0, 2, 'TOC')
ws.write(0, 3, 'Title')
ws.write(0, 4, 'Date')
ws.write(0, 5, 'DOI')


# crwaling information from journal_url file
journal_url = pd.read_excel(
    './journal_url.xlsx', sheet_name='url', names=['Abb', 'Link'])
journal_urls = journal_url['Link'].tolist()
Abb = journal_url['Abb'].tolist()

row = 1  # TOC row
for i in range(len(journal_urls)):
    link = journal_urls[i]
    abbreviation = Abb[i]
    req = requests.get(link)
    soup = BeautifulSoup(req.content, 'html.parser')

    for j in range(num):
        toc = soup.select('div.issue-item_img > img')[j]  # Load the TOC
        image_url = 'https://pubs.acs.org' + str(toc)[35:-3]
        print(image_url)
        res = requests.get(image_url)
        image_data = BytesIO(res.content)  # Process the image file
        ws.insert_image('C%d' % (row+1), image_url,
                        {'x_scale': 0.45, 'y_scale': 0.45, 'image_data': image_data})
        ws.write(row, 0, row)
        ws.write(row, 1, abbreviation)
        row += 1
print("---------------------------------------------")
print("----1/3 Save the TOC from ACS publication----")
print("---------------------------------------------")


workbook.close()

# crawling information from journal_url file
journal_url = pd.read_excel(
    './journal_url.xlsx', sheet_name='url', names=['Abb', 'Link'])
journal_urls = journal_url['Link'].tolist()
Abb = journal_url['Abb'].tolist()

# Empty list and df
df = pd.DataFrame(columns=['Abb', 'No.', 'Title', 'Date', 'DOI'])
summary = []
doi = []
data = []
count = 1  # Paper number of Journal

# ACS journal option
for i in range(len(journal_urls)):
    link = journal_urls[i]
    abbreviation = Abb[i]
    options = webdriver.ChromeOptions()  # hide chromedriver
    options.add_argument("headless")  # hide chromedriver
    driver = webdriver.Chrome(
        'chromedriver.exe', options=options)  # hide chromedriver
    driver.get(link)
    req = requests.get(link)
    soup = BeautifulSoup(req.content, 'html.parser')

    for j in range(num):  # Crwaling the ACS ASAP journal
        title = driver.find_elements_by_class_name('issue-item_title')[j]
        date = driver.find_elements_by_class_name('pub-date-value')[j]
        doi = soup.select(
            'div.issue-item_metadata > span > h5 > a')[j]['href']  # Load the doi
        summary.append([abbreviation, count, title.text,
                       date.text, "https://pubs.acs.org/doi"+doi[4:]])  # can modify sci-hub
        count += 1
    driver.close()
df = df.append(pd.DataFrame(summary, columns=[
               'Abb', 'No.', 'Title', 'Date', 'DOI']))

print(df)
print("---------------------------------------------")
print("---2/3 Save the Bibloigraphic Information----")
print("---------------------------------------------")
writer = pd.ExcelWriter('summary.xlsx', mode='a', engine='openpyxl')
df.to_excel(writer, sheet_name='Information', index=False)

writer.save()
print("---------------------------------------------")
print("-----------------3/3 Finish------------------")
print("---------------------------------------------")
