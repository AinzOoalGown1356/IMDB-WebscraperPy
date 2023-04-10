from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'TOP 250 MOVIES OAT'

sheet.append(['Movie Rank', 'Movie Name', 'Release Date', 'IMDB Rating'])

url = 'https://www.imdb.com/chart/top/'

try:
    source = requests.get(url)
    source.raise_for_status()

    soup = BeautifulSoup(source.text, 'html.parser')

    movies = soup.find('tbody', class_="lister-list").find_all('tr') 

    for movie in movies:
        name = movie.find('td', class_="titleColumn").a.text
        number = movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0]
        releasedate = movie.find('td', class_="titleColumn").span.text 
        imdbrating = movie.find('td', class_="ratingColumn imdbRating").strong.text 
        info = [number, name, releasedate, imdbrating]

    # Filter movies by changing the release date
    if "1990" in releasedate:
        sheet.append([number, name, releasedate, imdbrating])
        excel.save('TOP MOVIES 250 OAT.xlsx')
    
    # Filter movies by their IMDB rating
    if int(imdbrating) <= 7.5:
       sheet.append([number, name, releasedate, imdbrating])
       excel.save('TOP MOVIES 250 OAT.xlsx')

except requests.exceptions.HTTPError as http_error:
    print(f"HTTP error occurred: {http_error}")
except requests.exceptions.Timeout as timeout_error:
    print(f"Request timed out: {timeout_error}")
except requests.exceptions.ConnectionError as connection_error:
    print(f"Connection error occurred: {connection_error}")
except requests.exceptions.RequestException as error:
    print(f"An error occurred: {error}")
