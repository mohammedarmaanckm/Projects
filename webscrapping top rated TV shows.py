from bs4 import BeautifulSoup
import requests, openpyxl

excel=openpyxl.Workbook()
print(excel.sheetnames)
sheet=excel.active
sheet.title='Top Rated TV Shows'
sheet.append(['Show Rank','Show name','Year of Release','IMDB rating'])

try:
    source=requests.get('https://www.imdb.com/chart/toptv/')
    source.raise_for_status()
    soup=BeautifulSoup(source.text,'html.parser')
    movies=soup.find('tbody',class_='lister-list').find_all('tr')
    for x in movies:
        name=x.find('td',class_='titleColumn').a.text
        rank=x.find('td',class_='titleColumn').get_text(strip=True).split('.')[0]
        year=x.find('td',class_='titleColumn').span.text.strip('()')
        rating=x.find('td',class_='ratingColumn imdbRating').strong.text
        print(rank,name,year,rating)
        sheet.append([rank,name,year,rating])
except Exception as e:
    print(e)
excel.save("IMDB TV Shows ratings.xlsx")






