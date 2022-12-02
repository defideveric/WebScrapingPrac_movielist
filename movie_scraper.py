from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
sheet =excel.active
sheet.title ='Top Rated Movies'
print(excel.sheetnames)
sheet.append(['Movie Rank', 'Movie Name', 'Year of Release', 'IMBD Rating'])


source = requests.get("https://www.imdb.com/chart/top/")

soup = BeautifulSoup(source.text, 'html.parser')

movies = soup.find('tbody', class_="lister-list").find_all("tr")

for movie in movies:

    name = movie.find('td', class_='titleColumn').a.text

    rank = movie.find('td', class_='titleColumn').get_text(strip=True).split('.')[0]

    year = movie.find('td', class_='titleColumn').span.text.strip('()')

    rating = movie.find('td', class_='ratingColumn imdbRating').strong.text


    print(name, rank, year, rating)
    sheet.append([name, rank, year, rating])

excel.save('IMDB Moving Ratings.xlsx')