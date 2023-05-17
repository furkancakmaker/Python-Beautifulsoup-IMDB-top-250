import requests, openpyxl
from bs4 import BeautifulSoup

# Excel dosyası oluşturup sütün başlıklarını ekliyoruz
excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = "IMDB Top 250"
sheet.append(["Sıralama", "Film Adı", "Yıl", "IMDB Puanı"])
sheet.column_dimensions["B"].width = 45

# Siteye istek gönderip BeautifulSoup ile kodları ayrıştırıyoruz 
source = requests.get("https://www.imdb.com/chart/top/")
soup = BeautifulSoup(source.text, "html.parser")

# Filmlerin olduğu tabloyu seçiyoruz
movies = soup.find("tbody", class_="lister-list").find_all("tr")

# Filmlerin sıralamasını, adını, yılını ve ımdb puanını alıyoruz
for movie in movies:
    name = movie.find("td", class_="titleColumn").a.text
    rank = movie.find("td", class_="titleColumn").get_text(strip=True).split(".")[0]
    year = movie.find("td", class_="titleColumn").span.text.strip("()")
    rating = movie.find("td", class_="ratingColumn imdbRating").strong.text

    print(rank, name, year, rating)

    # Aldığımız bilgileri excel hücrelerine ekliyoruz
    sheet.append([rank, name, year, rating])

# Son olarak excel dosyamızı kaydediyoruz
excel.save("IMDB top 250.xlsx")




