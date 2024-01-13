import pandas as pd
import numpy as np

class Parser:

    def __init__(self, year):
        self.year = year

    def extract(self, path, sheet=1):
        self.df = pd.read_excel(path,sheet)
        print('Input file loaded into dataframe')

    def transform(self):

        column_mapping = {"Schengen State": "schengen_state"
                         ,"Country where consulate is located": "origin_country"
                         ,"Consulate": "origin_consulate"
                         ,"Uniform visas applied for": "visas_applied"
                         ,"Total  uniform visas issued (including MEV) \n": "visas_issued"
                         ,"Uniform visas not issued": "visas_not_issued"}
        if self.year == 2013:
            column_mapping = {"Schengen State": "schengen_state"
                        ,"Country where consulate is located": "origin_country"
                        ,"Consulate": "origin_consulate"
                        ,"C visas applied for": "visas_applied"
                        ,"Total C uniform visas issued (including MEV) \n": "visas_issued"
                        ,"C visas not issued": "visas_not_issued"}

        self.df = self.df.rename(columns = column_mapping)
        print('Renamed columns')

        self.df = self.df[self.df["schengen_state"].notna()]
        self.df = self.df[self.df["origin_country"].notna()]
        print('Removed irrelevant rows such as summary rows')

        self.df = self.df[["schengen_state", "origin_country","origin_consulate","visas_applied", "visas_issued", "visas_not_issued"]]
        print('Selected columns')

        self.df["visas_applied"].replace(np.nan, 0, inplace=True)
        self.df["visas_issued"].replace(np.nan, 0, inplace=True)
        self.df["visas_not_issued"].replace(np.nan, 0, inplace=True)
        print('Replaced applied, issued and not issued null values with 0')

        self.df["visas_issued"].loc[self.df["visas_issued"]<0] = 0
        print('Replaced negative entries for visas issued with 0')

        self.df["visa_refusal_rate"] = self.df["visas_not_issued"] / (self.df["visas_issued"] + self.df["visas_not_issued"])
        print('Calculate refusal rate as the not issued share of the total issued and not issued')

        self.df["year"] = self.year
        print('Added column with year issued and not issued')
    
    def load(self):
        self.df.to_csv('../data/silver/schengen-visa-' + str(self.year) + '.csv')