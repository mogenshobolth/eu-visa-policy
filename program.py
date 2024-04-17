import csv
import pandas as pd
import numpy as np
from dataclasses import dataclass, field

def clean_number(x) -> int:
    if pd.isna(x) or x < 0:
        return 0
    else:
        return int(x)

def map_country(x) -> str:
    return {
            'HONG KONG S.A.R.': 'CHINA',
            'CONGO (DEMOCRATIC REPUBLIC)': 'DEMOCRATIC REPUBLIC OF THE CONGO',
            'CONGO, THE DEMOCRATIC REPUBLIC OF THE': 'DEMOCRATIC REPUBLIC OF THE CONGO',
            'CONGO (BRAZZAVILLE)': 'REPUBLIC OF THE CONGO',
            'FORMER YUGOSLAV REPUBLIC OF MACEDONIA': 'MACEDONIA',
            'HOLY SEE': 'VATICAN',
            'HOLY SEE (VATICAN CITY STATE)': 'VATICAN',
            'LIBYAN ARAB JAMAHIRIYA': 'LIBYA',
            'TANZANIA, UNITED REPUBLIC OF': 'TANZANIA',
            'KOREA, REPUBLIC OF': 'SOUTH KOREA',
            'TAIWAN, PROVINCE OF CHINA': 'TAIWAN',
            'KOREA, DEMOCRATIC PEOPLE\'S REPUBLIC OF': 'NORTH KOREA',
            'MACAO S.A.R.': 'CHINA',
            'IRAN, ISLAMIC REPUBLIC OF': 'IRAN',
            'MOLDOVA, REPUBLIC OF': 'MOLDOVA',
            'MACAO': 'CHINA',
            'PUERTO RICO': 'USA'
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
        'SAN FRANCISCO, CA': 'SAN FRANCISCO'
    }.get(x, x)

@dataclass
class InputFile:
    """Class for keeping track of input files """
    
    file_name: str
    
    year: int

    sheet: str
    
    format: str

@dataclass
class InputFiles:
    """Class for keeping track of a collection of input files"""

    sheets: list[InputFile] = field(default_factory=list)

    def add_file_alpha(self, filename, year):
        self.sheets.append(InputFile(filename, year, 'Data for consulates', 'alpha'))

    def add_file_beta(self, filename, year):
        self.sheets.append(InputFile(filename, year, 'BG, CY, HR, RO', 'beta'))

    def add_file_gamma(self, filename, year):
        self.sheets.append(InputFile(filename, year, 'BG, HR, RO', 'gamma'))

    def add_file(self, filename, year, sheet, format):
        self.sheets.append(InputFile(filename, year, sheet, format))

@dataclass
class Line:
    """Class for keeping track of the content of a row in an input dataset """
    
    reporting_year: int
    
    reporting_state: str
    
    consulate_country: str
    
    consulate_city: str

    visitor_visa_applications: int

    visitor_visa_issued: int

    visitor_visa_not_issued: int

    def visitor_visa_refusal_rate(self) -> float:
        if self.visitor_visa_issued + self.visitor_visa_not_issued > 0:
            return self.visitor_visa_not_issued / (self.visitor_visa_issued + self.visitor_visa_not_issued)

    def toIterable(self):
        return iter(
            [
                self.reporting_year,
                self.reporting_state,
                self.consulate_country,
                self.consulate_city,
                self.visitor_visa_applications,
                self.visitor_visa_issued,
                self.visitor_visa_not_issued,
                self.visitor_visa_refusal_rate(),
            ]
        )

@dataclass
class Dataset:
    """Class for keeping track of a collection of lines"""

    line: list[Line] = field(default_factory=list)

    def to_header(self):
        return [
            "reporting_year",
            "reporting_state",
            "consulate_country",
            "consulate_city",
            "visitor_visa_applications",
            "visitor_visa_issued",
            "visitor_visa_not_issued",
            "visitor_visa_refusal_rate",
        ]
    
    def to_csv(self):
        with open('output/visitor-visa-statistics.csv', 'w') as f:
            write = csv.writer(f)
            write.writerow(self.to_header())
            for l in self.line:
                write.writerow(l.toIterable())

    def load(self, sheets):
        for s in sheets:
            print('Processing: ' + str(s.year))
            if s.file_name.split(".")[-1] == 'csv':
                df = pd.read_csv('input/' + s.file_name)
            else:
                df = pd.read_excel('input/' + s.file_name,s.sheet)                
            match s.format:
                case 'alpha':
                    df.apply(lambda x: self.add_line_alpha(s.year, x), axis=1)
                case 'beta':
                    df.apply(lambda x: self.add_line_beta(s.year, x), axis=1)
                case 'gamma':
                    df.apply(lambda x: self.add_line_gamma(s.year, x), axis=1)
                case 'delta':
                    df.apply(lambda x: self.add_line_delta(s.year, x), axis=1)
                case 'epsilon':
                    df.apply(lambda x: self.add_line_epsilon(x), axis=1)
                case _:
                    raise ValueError("Format not supported")

    def add_line_alpha(self, reporting_year, x):
        if pd.isna(x['Schengen State']) or pd.isna(x['Country where consulate is located']):
            return
        self.line.append(Line(reporting_year
            ,str(x['Schengen State']).strip()
            ,map_country(str(x['Country where consulate is located']).strip())
            ,map_city(str(x['Consulate']).strip())
            ,clean_number(x['Uniform visas applied for'])
            ,clean_number(x['Total  uniform visas issued (including MEV) \n'])
            ,clean_number(x['Uniform visas not issued'])))

    def add_line_beta(self, reporting_year, x):
        if pd.isna(x['Member State']) or pd.isna(x['Country where consulate is located']):
            return
        self.line.append(Line(reporting_year
            ,str(x['Member State']).strip()
            ,map_country(str(x['Country where consulate is located']).strip())
            ,map_city(str(x['Consulate']).strip())
            ,clean_number(x['Short-stay visas applied for'])
            ,clean_number(x['Total  short-stay visas issued (including MEV) \n'])
            ,clean_number(x['Short-stay visas not issued'])))

    def add_line_gamma(self, reporting_year, x):
        if pd.isna(x['Schengen State']) or pd.isna(x['Country where consulate is located']):
            return
        self.line.append(Line(reporting_year
            ,str(x['Schengen State']).strip()
            ,map_country(str(x['Country where consulate is located']).strip())
            ,map_city(str(x['Consulate']).strip())
            ,clean_number(x['Short-stay visas applied for'])
            ,clean_number(x['Total  short-stay visas issued (including MEV) \n'])
            ,clean_number(x['Short-stay visas not issued'])))

    def add_line_delta(self, reporting_year, x):
        if pd.isna(x['Schengen State']) or pd.isna(x['Country where consulate is located']):
            return
        self.line.append(Line(reporting_year
            ,str(x['Schengen State']).strip()
            ,map_country(str(x['Country where consulate is located']).strip())
            ,map_city(str(x['Consulate']).strip())
            ,clean_number(x['C visas applied for'])
            ,clean_number(x['Total C uniform visas issued (including MEV) \n'])
            ,clean_number(x['C visas not issued'])))

    def add_line_epsilon(self, x):
        self.line.append(Line(x['dYear']
            ,str(x['receivingCountryName']).strip()
            ,map_country(str(x['sendingCountryName']).strip())
            ,map_city(str(x['sendingCityName']).strip())
            ,clean_number(x['appliedABC'])
            ,clean_number(x['issuedABC'])
            ,clean_number(x['notIssuedABC'])))

inputfiles = InputFiles()
inputfiles.add_file_alpha('Visa statistics for consulates in 2022_en.xlsx',2022)
inputfiles.add_file_beta('Visa statistics for consulates in 2022_en.xlsx',2022)
inputfiles.add_file_alpha('Visa statistics for consulates 2021_0.xlsx',2021)
inputfiles.add_file_beta('Visa statistics for consulates 2021_0.xlsx',2021)
inputfiles.add_file_alpha('visa_statistics_for_consulates_2020.xlsx',2020)
inputfiles.add_file_beta('visa_statistics_for_consulates_2020.xlsx',2020)
inputfiles.add_file_alpha('2019-consulates-schengen-visa-stats.xlsx',2019)
inputfiles.add_file_beta('2019-consulates-schengen-visa-stats.xlsx',2019)
inputfiles.add_file_alpha('2018-consulates-schengen-visa-stats.xlsx',2018)
inputfiles.add_file_gamma('2018-consulates-schengen-visa-stats.xlsx',2018)
inputfiles.add_file_alpha('2017-consulates-schengen-visa-stats.xlsx',2017)
inputfiles.add_file_gamma('2017-consulates-schengen-visa-stats.xlsx',2017)
inputfiles.add_file_alpha('2016_consulates_schengen_visa_stats_en.xlsx',2016)
inputfiles.add_file_gamma('2016_consulates_schengen_visa_stats_en.xlsx',2016)
inputfiles.add_file_alpha('2015_consulates_schengen_visa_stats_en.xlsx',2015)
inputfiles.add_file_gamma('2015_consulates_schengen_visa_stats_en.xlsx',2015)
inputfiles.add_file_alpha('2014_global_schengen_visa_stats_compilation_consulates_-_final_en.xlsx',2014)
inputfiles.add_file('2014_global_schengen_visa_stats_compilation_consulates_-_final_en.xlsx',2014,'BG, HR, RO', 'alpha')
inputfiles.add_file('synthese_2013_with_filters_en.xls',2013,'Complete data', 'delta')
inputfiles.add_file('evd_visa_practice_eu.csv',None,None,'epsilon')

d = Dataset()
d.load(inputfiles.sheets)
d.to_csv()