import csv
import pandas as pd
import numpy as np
from dataclasses import dataclass, field

def missing_to_zero(x) -> int:
    return 0 if pd.isna(x) else int(x)

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
    
    def add_line_alpha(self, reporting_year, x):
        if pd.isna(x['Schengen State']):
            return
        self.line.append(Line(reporting_year
            ,str(x['Schengen State']).strip()
            ,str(x['Country where consulate is located']).strip()
            ,str(x['Consulate']).strip()
            ,missing_to_zero(x['Uniform visas applied for'])
            ,missing_to_zero(x['Total  uniform visas issued (including MEV) \n'])
            ,missing_to_zero(x['Uniform visas not issued'])))

    def add_line_beta(self, reporting_year, x):
        if pd.isna(x['Member State']):
            return
        self.line.append(Line(reporting_year
            ,str(x['Member State']).strip()
            ,str(x['Country where consulate is located']).strip()
            ,str(x['Consulate']).strip()
            ,missing_to_zero(x['Short-stay visas applied for'])
            ,missing_to_zero(x['Total  short-stay visas issued (including MEV) \n'])
            ,missing_to_zero(x['Short-stay visas not issued'])))

d = Dataset()

df = pd.read_excel('input/Visa statistics for consulates in 2022_en.xlsx','Data for consulates')
df.apply(lambda x: d.add_line_alpha(2022, x), axis=1)

df = pd.read_excel('input/Visa statistics for consulates in 2022_en.xlsx','BG, CY, HR, RO')
df.apply(lambda x: d.add_line_beta(2022, x), axis=1)

d.to_csv()