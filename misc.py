from urllib.request import urlopen
from json import loads
import pandas as pd

# TODO: if the amounts are in this format: 9,999.999 - change this to: 9999.999, otherwise Excel will read it as 9.999999? - check
# TODO: add try/except and manual workaround to below code in case of API fail, change in currency name, etc.
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