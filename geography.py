import pandas as pd
from dataclasses import dataclass, field

def map_country(x) -> str:
    return {
            'CZECH REPUBLIC': 'CZECHIA',
            'CONGO (DEMOCRATIC REPUBLIC)': 'CONGO, DEM. REP.',
            'CONGO, THE DEMOCRATIC REPUBLIC OF THE': 'CONGO, DEM. REP.',
            'CONGO (DEMOCRATIC REPUBLIC OF)': 'CONGO, DEM. REP.',
            'DEMOCRATIC REPUBLIC OF THE CONGO': 'CONGO, DEM. REP.',
            'CONGO (BRAZZAVILLE)': 'CONGO, REP.',
            'CONGO': 'CONGO, REP.',
            'REPUBLIC OF THE CONGO': 'CONGO, REP.',
            'HOLY SEE': 'ITALY',
            'HOLY SEE (VATICAN CITY STATE)': 'ITALY',
            'VATICAN': 'ITALY',
            'LIBYAN ARAB JAMAHIRIYA': 'LIBYA',
            'TANZANIA, UNITED REPUBLIC OF': 'TANZANIA',
            'KOREA, REPUBLIC OF': 'KOREA, REP.',
            'KOREA (SOUTH)': 'KOREA, REP.',
            'SOUTH KOREA': 'KOREA, REP.',
            'KOREA, DEMOCRATIC PEOPLE\'S REPUBLIC OF': 'KOREA, DEM. PEOPLE\'S REP.',
            'KOREA (NORTH)': 'KOREA, DEM. PEOPLE\'S REP.',
            'NORTH KOREA': 'KOREA, DEM. PEOPLE\'S REP.',
            'TAIWAN, PROVINCE OF CHINA': 'TAIWAN, CHINA',
            'TAIWAN': 'TAIWAN, CHINA',
            'MACAO S.A.R.': 'CHINA',
            'HONG KONG S.A.R.': 'CHINA',
            'MACAO': 'CHINA',
            'MACAO SAR': 'CHINA',
            'HONG KONG SAR': 'CHINA',
            'MOLDOVA, REPUBLIC OF': 'MOLDOVA',
            'BRUNEI': 'BRUNEI DARUSSALAM',
            'CÔTE D\'IVOIRE': 'CÔTE D’IVOIRE',
            'COTE D\'IVOIRE': 'CÔTE D’IVOIRE',
            'LAO PEOPLE\'S DEMOCRATIC REPUBLIC': 'LAO PDR',
            'LAOS': 'LAO PDR',
            'NERTHERLANDS': 'NETHERLANDS',
            'MACEDONIA': 'NORTH MACEDONIA',
            'FORMER YUGOSLAV REPUBLIC OF MACEDONIA': 'NORTH MACEDONIA',
            'RUSSIA': 'RUSSIAN FEDERATION',
            'TURKEY': 'TÜRKIYE',
            'VIET NAM': 'VIETNAM',
            'SYRIA': 'SYRIAN ARAB REPUBLIC',
            'USA': 'UNITED STATES',
            'UNITED STATES OF AMERICA': 'UNITED STATES',
            'EGYPT': 'EGYPT, ARAB REP.',
            'IRAN': 'IRAN, ISLAMIC REP.',
            'IRAN, ISLAMIC REPUBLIC OF': 'IRAN, ISLAMIC REP.',
            'SLOVAKIA': 'SLOVAK REPUBLIC',
            'SAINT LUCIA': 'ST. LUCIA',
            'VENEZUELA': 'VENEZUELA, RB',
            'KYRGYZSTAN': 'KYRGYZ REPUBLIC',
            'PALESTINIAN AUTHORITY': 'WEST BANK AND GAZA',
            'PALESTINE': 'WEST BANK AND GAZA',
            'CAPE VERDE': 'CABO VERDE',
            'SAO TOME AND PRINCIPE': 'SÃO TOMÉ AND PRÍNCIPE',
            'YEMEN': 'YEMEN, REP.',
            'BURMA': 'MYANMAR',
            'AUSTRILIA': 'AUSTRALIA'
        }.get(x, x)

def map_city(x) -> str:
    return {
        'YAONDE': 'YAOUNDE',
        'VITSYEBSK': 'VITEBSK',
        'BANDAR SERI BEGWAN': 'BANDAR SERI BEGAWAN',
        'GHIROKASTER': 'GJIROKASTER',
        'CIDADE DA PRAIA': 'PRAIA',
        'MIAMI, FL': 'MIAMI',
        'NEW YORK, NY': 'NEW YORK',
        'SAN FRANCISCO': 'SAN FRANCISCO',
        'CHICAGO, IL': 'CHICAGO',
        'HOUSTON, TX': 'HOUSTON',
        'LOS ANGELES, CA': 'LOS ANGELES',
        'BOSTON, MA': 'BOSTON',
        'DETROIT, MI': 'DETROIT',
        'NEWARK, NJ': 'NEWARK',
        'TAMPA, FL': 'TAMPA',
        'NEW BEDFORD, MA': 'NEW BEDFORD',
        'CLEVELAND, OH': 'CLEVELAND',
        'VINNYTSYA': 'VINNYTSIA',
        'WILLEMSTAD (CURACAO)': 'WILLEMSTAD',
        'BELEM, PA': 'BELEM',
        'SAN FRANCISCO, CA': 'SAN FRANCISCO',
        'KABUl': 'KABUL',
        'ANDORRA-LA-VELLA': 'ANDORRA LA VELLA',
        'ROSARIO (Santa Fé)': 'ROSARIO - SANTA FE',
        'BELÉM': 'BELEM',
        'SALVADOR-BAHIA': 'SALVADOR DE BAHIA',
        'SANTIAGO DE CHILE': 'SANTIAGO',
        'GUANGZHOU (CANTON)': 'GUANGZHOU',
        'ADDIS ABEBA': 'ADDIS ABABA',
        'TBILISSI': 'TBILISI',
        'PORT-AU-PRINCE': 'PORT AU PRINCE',
        'BÉKÉSCSABA': 'BEKESCSABA',
        'BAGDAD': 'BAGHDAD',
        'OSAKA-KOBE': 'OSAKA',
        'KUWAIT': 'KUWAIT CITY',
        'LUXEMBURG': 'LUXEMBOURG',
        'FES': 'FEZ',
        'MARRAKECH': 'MARRAKESH',
        'TANGER': 'TANGIER',
        'POINTE-NOIRE': 'POINTE NOIRE',
        'TIMIȘOARA': 'TIMISOARA',
        'NOVOROSSIISK': 'NOVOROSSIYSK',
        'NOVOROSSISK': 'NOVOROSSIYSK',
        'ST. PETERSBURG': 'ST PETERSBURG',
        'JEDDA': 'JEDDAH',
        'BELGRAD': 'BELGRADE',
        'CAPETOWN': 'CAPE TOWN',
        'VALENCIA (SPAIN)': 'VALENCIA',
        'DAR-ES-SALAAM': 'DAR ES SALAAM',
        'PORT-OF-SPAIN': 'PORT OF SPAIN',
        'CHERNIVITSI': 'CHERNIVTSI',
        'KIEV': 'KYIV',
        'LVOV': 'LVIV',
        'ODESSA': 'ODESA',
        'SEBASTOPOL': 'SEVASTOPOL',
        'VINNITSA': 'VINNYTSIA',
        'ATLANTA, GA': 'ATLANTA',
        'PHILADELPHIA, PA': 'PHILADELPHIA',
        'SAN JUAN (PORT RICO)': 'SAN JUAN',
        'SAN JUAN, PR': 'SAN JUAN',
        'WASHINGTON, DC': 'WASHINGTON',
        'VATICAN CITY (ROME)': 'VATICAN CITY',
        'HO CHI MINH': 'HO-CHI MINH CITY',
        'SANA\'A': 'SANAA',
        'SANA \'A': 'SANAA'
    }.get(x, x)

@dataclass
class Country:

    code: str

    name: str

    region: str

    income_group: str

@dataclass
class CountryList:

    filepath: str = 'input/worldbank-classification-2023.xlsx'

    countries: list[Country] = field(default_factory=list)

    def load(self):
        df = pd.read_excel(self.filepath,0)
        df.apply(lambda x: self.add_country(x), axis=1)

    def add_country(self, x): 
        if pd.isna(x['Economy']) or pd.isna(x['Code']) or pd.isna(x['Region']):
            return
        income_group = str(x['Income group'])
        if str(x['Income group']) == 'Venezuela, RB': 
            income_group = 'Upper middle income'
        self.countries.append(Country(str(x['Code']), str(x['Economy']), str(x['Region']), income_group))

    def get(self, name):
        name = name.strip().upper()
        name = map_country(name)
        try:
            c = next(x for x in self.countries if x.name.upper() == name)
            return c
        except StopIteration:
            raise ValueError('Not found: ' + name)