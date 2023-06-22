import utm 
from geopy.geocoders import Nominatim
import requests
from pandas import *
global zoneVal
from ambiance import Atmosphere 

def locator(coordinates):
    try:
        geolocator = Nominatim(user_agent = "geoapiExercises")
        city = geolocator.reverse(coordinates)
        print('adding city')
    except Exception as e:
        pass

    if city.raw['address'].get('province', '')=='':
        if city.raw['address'].get('state_district', '') != '':
            province = city.raw['address'].get('state_district', '')
        else: province = city.raw['address'].get('state', '')
    else: province = city.raw['address'].get('province', '')


    if province =='Alacant / Alicante':
        province = 'Alicante'
    elif province == 'Araba/Álava':
        province = 'Álava'
    elif province == 'Asturias / Asturies':
        province = 'Asturias'
    elif province == 'Comunitat Valenciana':
        province = 'Castellón'
    elif province == 'Comunidad de Madrid':
        province = 'Madrid'
    elif province == 'Región de Murcia':
        province = 'Murcia'
    elif province == 'Navarra - Nafarroa':
        province = 'Navarra'
    elif province == 'València / Valencia':
        province = 'Valencia'
    else:
        pass
    return province

def latLonToXY(lat,lon):
    coordinatesUTM = utm.from_latlon(lat,lon)
    return coordinatesUTM #UTMx, UTMy, huso

def get_elevationAndPressure(lat, lon):
    url = ('https://api.opentopodata.org/v1/test-dataset'f'?locations={lat},{lon}')
    while True:
        try:
           response = requests.get(url).json()
        except Exception as e:
           continue
        break
   
    elevation = response['results'][0]['elevation']
    pressure = (Atmosphere(elevation).pressure[0])*(76/101325) #in cmHg

    return elevation, pressure 

pass

