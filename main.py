import requests
from bs4 import BeautifulSoup
import openpyxl
import re

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title="Movies"
sheet.append(["Rank","MovieName","Year","Ratings"])
try:

    req = requests.get("https://www.imdb.com/chart/top/?ref_=nv_mv_250")
    soup = BeautifulSoup(req.content,"html.parser")
    main = soup.find("tbody",class_="lister-list").find_all("tr")
    #print(main)
    for i in main:
        Rank = i.find("td",class_="titleColumn").get_text(strip=True).split(".")[0]
        MovieName = i.find("td",class_="titleColumn").a.text
        Years = i.find("td",class_="titleColumn").span.text
        Year = re.sub("\W","",Years)
        Ratings = i.find("td",class_="ratingColumn imdbRating").get_text(strip=True)
        #print(Rank,MovieName,Year,Ratings)
        sheet.append([Rank,MovieName,Year,Ratings])
except:
    print("Error Occured")
excel.save("IMDB.xlsx")
