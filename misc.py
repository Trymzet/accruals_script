from urllib.request import urlopen
from json import loads

# TODO: if the amounts are in this format: 9,999.999 - change this to: 9999.999, otherwise Excel will read it as 9.999999? - check
# TODO: add try/except and manual workaround to below code in case of API fail, change in currency name, etc.
# TODO: add a workaround for Costa Rica (CRC)
"""import pandas as pd
from subprocess import call
from os import remove
import numpy as np
from locale import setlocale, LC_NUMERIC, atof

setlocale(LC_NUMERIC, 'English_Canada.1252')

csv = "EXP031-RPT-Process-Accruals_with_Expense_Report.csv"

output = pd.read_csv(csv, skiprows=[0], encoding="latin-1", engine="python", na_values="")
print(output.head())
print(output.dtypes)
print(float(atof("15,555.00")))

output["Net Amount LC"] = output["Net Amount LC"].apply(lambda x: x.replace(",", "") if type(x) != float else x)
print(output["Net Amount LC"].tail())
print(output["Net Amount LC"].astype(float).head())"""

WD_report_name = "EXP031-RPT-Process-Accruals_with_Expense_Report.xlsx"
# print("{}{}".format(WD_report_name[:-5], ".csv"))

# https://www.exchangerate-api.com API key - 1000 calls/month
exchange_rates_API_key = "88fcc0b255987f5639edd24f"
api_call = "https://v3.exchangerate-api.com/bulk/{}/USD".format(exchange_rates_API_key)
rates_api_response = urlopen(api_call)
rates_api_response_str = rates_api_response.read().decode("ascii")
rates_api_response_dict = loads(rates_api_response_str)
rates = rates_api_response_dict["rates"]
print(1/rates["CZK"])


# feel free to update company codes/currencies
currencies_in_scope = {"AUSA": "AUD", "BESA": "EUR", "BGSA": "BGN", "BRSA": "BRL", "CASA": "CAD", "CHSD": "CHF", "CNSA": "CNY",
                       "CRSB": "CRC", "CZSA": "CZK", "DESA": "EUR", "DKSA": "DKK", "ESSA": "EUR", "FRSA": "EUR", "GBF0": "USD",
                       "GBSA": "GBP", "IESA": "EUR", "IESB": "EUR", "ILSA": "ILS", "ILSB": "ILS", "INSA": "INR", "INSB": "INR",
                       "INSD": "INR", "ITSA": "EUR", "JPSA": "JPY", "LUSB": "EUR", "MXSC": "MXN", "NLSC": "EUR", "PHSB": "PHP",
                       "PLSA": "PLN", "PRSA": "PYG", "ROSA": "RON", "RUSA": "RUB", "SESA": "SEK", "TRSA": "TRY", "USMS": "USD",
                       "USSM": "USD", "USSN": "USD"}
exchange_rates_to_USD = {}
for company_code in currencies_in_scope:
    currency = currencies_in_scope[company_code]
    # the rates from API are from USD to x; we need from x to USD
    try:
        exchange_rate_to_USD = 1/rates[currency]
    except:
        continue
    if company_code in exchange_rates_to_USD:
        continue
    else:
        exchange_rates_to_USD[company_code] = exchange_rate_to_USD

print(exchange_rates_to_USD["CZSA"])

"""
Now simply check if e.g. the sum in column "CREDIT" times exchange rate for the company code's currency is > 5000 -- in pivot.py
"""