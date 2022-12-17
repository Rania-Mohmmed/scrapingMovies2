import openpyxl
from bs4 import BeautifulSoup
import  requests
excel=openpyxl.Workbook()
sheet =excel.active
sheet.title='top Rated movies'
print(excel.sheetnames)
sheet.append(['Rank Movie','Movie Name','Year Of Releas','Ratint IMDb'])

try:
  source=requests.get('https://www.imdb.com/chart/top/')
  source.raise_for_status()
  soup=BeautifulSoup(source.text,'html.parser')
  movies=soup.find('tbody',{'class','lister-list'}).findAll('tr')
  for movie in movies:
      name=movie.find('td',{'class','titleColumn'}).a.text
      rank= (movie.find('td', {'class', 'titleColumn'}).text).strip().split('.')[0]#????['1', '\n      The Shawshank Redemption\n(1994)']
      year=movie.find('td', {'class', 'titleColumn'}).span.text.strip('()')
      ratting=movie.find('td',{'class','ratingColumn imdbRating'}).strong.text

      sheet.append([rank,name,year,ratting])

except Exception as e:
    print(e)
excel.save('TOP Movie Rating.xlsx')