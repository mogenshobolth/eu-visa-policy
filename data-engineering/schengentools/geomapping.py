import numpy as np
import pandas as pd
from geopy.geocoders import Nominatim

class Geocoder:
        
    def __init__(self):
        self.geolocator = Nominatim(user_agent="mhobolth-geo")

    def __lookup_location(self, country, city):
        try:
            point = self.geolocator.geocode(city + ', ' + country).point
            return pd.Series({'latitude': point.latitude, 'longitude': point.longitude})
        except:
            return None

    def __lookup_city(self, city):
        return {
            'YAONDE': 'YAOUNDE',
            'VITSYEBSK': 'VITEBSK',
            'BANDAR SERI BEGWAN': 'BANDAR SERI BEGAWAN',
            'GHIROKASTER': 'GJIROKASTER'
        }.get(city, city)

    def __lookup_country(self, country):
        return {
            'HONG KONG S.A.R.': 'HONG KONG',
            'CONGO (DEMOCRATIC REPUBLIC)': 'DEMOCRATIC REPUBLIC OF THE CONGO',
            'CONGO (BRAZZAVILLE)': 'REPUBLIC OF THE CONGO',
            'FORMER YUGOSLAV REPUBLIC OF MACEDONIA': 'MACEDONIA',
            'HOLY SEE': 'VATICAN'
        }.get(country, country)

    def read_schengen_csv(self, path):
        self.df = pd.read_csv(path)
        print('Input file loaded into dataframe')

    def extract_locations(self):
        self.df = self.df[["origin_country","origin_consulate"]]
        self.df["country"] = self.df["origin_country"]
        self.df["city"] = self.df["origin_consulate"]
        self.df = self.df.drop_duplicates()
        print('Extracted countries and cities from input dataset')

    def rename_cities(self):
        self.df['city'] = self.df.apply(lambda x: self.__lookup_city(x['city']), axis=1)
        print('Renamed relevant city names to geolocator compatiable names')

    def rename_countries(self):
        self.df['country'] = self.df.apply(lambda x: self.__lookup_country(x['country']), axis=1)
        print('Renamed relevant country names to geolocator compatiable names')

    def geocode_dataframe(self):
        self.df[['latitude', 'longitude']] = self.df.apply(lambda x: self.__lookup_location(x['country'], x['city']), axis=1)
        print('{}% of cities were geocoded!'.format((1 - sum(np.isnan(self.df['latitude'])) / len(self.df)) * 100))
    
    def save(self, path):
        self.df.to_csv(path)
        print('Output file saved')