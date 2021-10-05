from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
from bs4 import BeautifulSoup
import requests

# crwaling information from journal_url file
journal_url = pd.read_excel(
    './journal_url.xlsx', sheet_name='url', names=['Abb', 'Link'])
journal_urls = journal_url['Link'].tolist()
Abb = journal_url['Abb'].tolist()

# Empty list and df
df = pd.DataFrame(columns=['Abb', 'No.', 'Title', 'Date', 'DOI'])
summary = []
toc = []
doi = []
data = []
pdf = []
count = 1 

num = int(input("Please enter the number of journals : ", ))

# ACS journal option
for i in range(len(journal_urls)):
    link = journal_urls[i]
    abbreviation = Abb[i]
    driver = webdriver.Chrome('chromedriver.exe')
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

with pd.ExcelWriter("./summary.xlsx") as excel_writer:
    df.to_excel(excel_writer, sheet_name='ASAP', index=False)
