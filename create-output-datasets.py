import pandas as pd

df = pd.read_csv('output/visitor-visa-statistics.csv')
#df2 = df.groupby(['reporting_year','reporting_state'])['visitor_visa_applications'].sum()
#df2.to_csv('output/visitor-visa-statistics-year-state.csv')

df2 = df['reporting_state'].drop_duplicates()
df2.to_csv('output/visitor-visa-statistics-reporting-states.csv', index=False)

df2 = df[['consulate_country','consulate_city']].drop_duplicates()
df2.to_csv('output/visitor-visa-statistics-consulates.csv', index=False)

df2 = df[['consulate_country']].drop_duplicates()
df2.to_csv('output/visitor-visa-statistics-consulate-countries.csv', index=False)

#df2 = df[['reporting_year','reporting_state','visitor_visa_applications']]
#df2 = df2.groupby(['reporting_year','reporting_state'])['visitor_visa_applications'].sum()
#df2 = df2.pivot(columns='reporting_year',values='visitor_visa_applications')

df.pivot_table(
    values="visitor_visa_applications", index="reporting_state", columns="reporting_year", aggfunc="sum"
).to_csv('output/visitor-visa-statistics-year-state.csv')

df.groupby(['reporting_year'])['visitor_visa_applications'].sum().to_csv('output/visitor-visa-statistics-years.csv')

df.groupby(['reporting_year'])['consulate_city'].count().to_csv('output/visitor-visa-statistics-consulates.csv')

df.groupby(['reporting_year'])['visitor_visa_refusal_rate'].mean().to_csv('output/visitor-visa-statistics-consulate-refusal-rate.csv')

df.groupby(['reporting_year']).agg({'visitor_visa_issued':'sum','visitor_visa_not_issued':'sum'}).to_csv('output/visitor-visa-statistics-years-issued.csv')