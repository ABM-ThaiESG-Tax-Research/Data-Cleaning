import requests
import pandas as pd

respone = requests.get('https://dataservices.mof.go.th/api/data/income/get_data_overview?from=2004-01-01&to=2024-12-30')

df = pd.DataFrame(respone.json()['items'])

clean = df[['month_year','rd_pnd','month_year_th']]

with pd.ExcelWriter("D:/งาน/SET/tax_data_personal.xlsx") as writer:
    clean.to_excel(writer,index=False)