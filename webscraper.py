import requests
import time
import re
from bs4 import BeautifulSoup
from openpyxl import Workbook

book = Workbook()
sheet = book.active
timestr = time.strftime("%d%m%Y_%H"+"h"+"%M")

vehicle = []

URLS = open("urls.txt", "r")

for URL in URLS :

    URL = URL.replace("\n", "")

    page = requests.get(URL)
    html = BeautifulSoup(page.text, 'html.parser')


    brand = html.find(class_ = 'name').text.strip()
    
    model_check = html.find(class_ = 'h4 title fw-light mb-3')

    if str(type(model_check)) == "<class 'NoneType'>" :
        print(URL + " is invalid. Please check it")
        continue

    model = html.find(class_ = 'h4 title fw-light mb-3').text.strip()

    myfuels = html.find_all("h4", {"class": "fuel-title"})

    for fuel in myfuels :
        fuel_div = fuel.find_parent("div")
        
        fuel_value = fuel.text.strip()

        version_list = fuel_div.find_all("a", {"class": "car-name"})


        for version in version_list:
            version_value = version.text.strip()
            
            car_div = version.find_parent("div").find_parent("div")

            price = car_div.find_all("span")[0].text.strip()
            price = price.replace("€", "")
           #price = price.replace(".", ",") -> pour éviter d'avoir des prix qui ne sont ensuite pas formatés correctement 
            price = re.sub(r"\s+", "", price, flags=re.UNICODE)
            """
            Dans Excel, sélectionner la colonne des prix, cliquer sur le '!' puis 'convertir en nombre'
            afin de transformer l'ensemble de la colonne en valeur monétaire exploitable
            """                     

            tmcBXWL = car_div.find_all("li", {"class": "bx"})[1].text.strip()
            tmcBXWL = tmcBXWL.replace("Taxe de mise en circulation: ", "")
            tmcVL = car_div.find_all("li", {"class": "vl"})[1].text.strip()
            tmcVL = tmcVL.replace("Taxe de mise en circulation: ", "")

            taBXWL = car_div.find_all("li", {"class": "bx"})[2].text.strip()
            taBXWL = taBXWL.replace("Taxe annuelle: ", "")
            taVL = car_div.find_all("li", {"class": "vl"})[2].text.strip()
            taVL = taVL.replace("Taxe annuelle: ", "")


            vehicle = [brand, model, fuel_value, version_value, price, tmcBXWL, taBXWL, tmcVL, taVL]
            #print(vehicle)

            sheet.append(vehicle)

book.save(timestr + '.xlsx')

"""
Ajouter forme et style aux lignes & colonnes du tableau : largeur auto des cellules, création auto d'un tableau?
Raccourci pour sélectionner automatiquement toutes les données d'une feuille : CTRL+MAJ+FIN
"""
