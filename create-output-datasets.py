import pandas as pd

df = pd.read_csv('output/visitor-visa-statistics.csv')

# Summary statistics on applications
df.pivot_table(
    values="visitor_visa_applications", index="reporting_year", columns="consulate_country_region", aggfunc="sum"
).to_csv('output/visitor-visa-statistics-applications-sum-year-region.csv')

# Avg. refusal rate by consulate country and year over last five years, for geo-mapping 
df[df['reporting_year'] >= 2018].groupby(['consulate_country','consulate_country_code'])['visitor_visa_refusal_rate'].mean().to_csv('output/visitor-visa-statistics-2018to2022-country-refusal-rate.csv')

# Avg. refusal rate by region and year
df.pivot_table(
    values="visitor_visa_refusal_rate", index="reporting_year", columns="consulate_country_region", aggfunc="mean"
).to_csv('output/visitor-visa-statistics-refusal-rate-mean-year-region.csv')

# Avg. refusal rate by income group and year
df.pivot_table(
    values="visitor_visa_refusal_rate", index="reporting_year", columns="consulate_country_income_group", aggfunc="mean"
).to_csv('output/visitor-visa-statistics-refusal-rate-year-mean-income-group.csv')

# Variation in refusal rate by region and year
df.pivot_table(
    values="visitor_visa_refusal_rate", index="reporting_year", columns="consulate_country_region", aggfunc="std"
).to_csv('output/visitor-visa-statistics-refusal-rate-stddev-year-region.csv')

# Variation in refusal rate by income group and year
df.pivot_table(
    values="visitor_visa_refusal_rate", index="reporting_year", columns="consulate_country_income_group", aggfunc="std"
).to_csv('output/visitor-visa-statistics-refusal-rate-stddev-year-income-group.csv')



