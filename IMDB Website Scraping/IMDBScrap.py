from bs4 import BeautifulSoup
import requests
import pandas as pd
import openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Top Rated Movies'
sheet.append(['Movie Rank',
              'Movie Name',
              'Released Year',
              'IMDB Rating'])

try:
    source = requests.get('https://www.imdb.com/chart/top/')
    #print(source)
    source.raise_for_status()
    
    soup = BeautifulSoup(source.text, 'html.parser')
    
    movies = soup.find('tbody', class_ = 'lister-list').find_all('tr')
    #print(len(movies))
    
    for movie in movies:
        
        Rank = movie.find('td', class_='titleColumn').get_text(strip=True).split(".")[0]
        Name = movie.find('td', class_='titleColumn').a.text
        Year = movie.find('td', class_='titleColumn').span.text.strip('()')
        Rating = movie.find('td', class_='ratingColumn imdbRating').strong.text
        
        # output = {
        #     'Rank': Rank,
        #     'Movie_Name': Name,
        #     'Released_Year': Year,
        #     'Rating': Rating
        # }
        
        sheet.append([Rank, Name, Year, Rating])
        
except Exception as e:
    print(e)
    
excel.save('IMDB Movie Rating.xlsx')