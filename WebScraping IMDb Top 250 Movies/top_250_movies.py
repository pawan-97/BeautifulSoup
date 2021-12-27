# Importing required libraries
from bs4 import BeautifulSoup
import requests, openpyxl

#Creating excel file
excel = openpyxl.Workbook()
sheet = excel.active
#Changing sheet name
sheet.title = 'IMDb Top 250 Movies'
#Adding a row
sheet.append(['Rank','Movie Name','Year of release','IMDB Rating'])

# Scraping required data
try:
    html = requests.get('https://www.imdb.com/chart/top/')
    html.raise_for_status()

    soup = BeautifulSoup(html.text,'html.parser')
    movies = soup.find('tbody',class_='lister-list').find_all('tr')
    for movie in movies:
        name = movie.find('td',class_='titleColumn').a.text
        rank = movie.find('td',class_='titleColumn').get_text(strip=True).split('.')[0]
        year = movie.find('td',class_='titleColumn').span.text.strip('()')
        rating = movie.find('td',class_='ratingColumn imdbRating').strong.text
        print(rank,name, year, rating)
        sheet.append([rank,name, year, rating])
except Exception as e:
    print(e)

#Saving the excel file
excel.save('IMDb Top 250 Movies.xlsx')
