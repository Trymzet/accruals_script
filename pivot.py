import pandas as pd
import openpyxl as oxl

clean = pd.read_excel("clean.xlsx")

#pivot = {}

#for line in clean:
#    checksums={}


cleaner = clean[["Entity Code", "Checksum", "Acc#", "Expense Report Number", "Net Amount LC"]]

#pivot1 = pd.pivot_table(clean[cols_to_pivot[0]], values=amount_col, rows=cols_to_pivot)
pivot = cleaner.groupby(["Entity Code", "Checksum", "Expense Report Number"], as_index=False)

co_codes = []
for key, group in pivot:
    #print(key[0])
    #print(group)
    co_code = key[0]
    if co_code not in co_codes:
        co_codes.append(co_code)

JE_csv_columns = ["ACCOUNT", "DEBIT", "CREDIT", "TAX CODE", "LINE MEMO", "MRU", "BUSINESS AREA", "PROFIT CENTER", "FUNCTIONAL AREA",
                  "DATE", "ACCOUNTING BOOK", "SUBSIDIARY", "CURRENCY", "MEMO", "REVERSAL DATE", "TO SUBSIDIARY", "UNIQUE ID"]
JE_csv = pd.DataFrame(columns=JE_csv_columns)

last_day_of_previous_month = (pd.to_datetime("today") - pd.tseries.offsets.MonthEnd(1))#.strftime("%m/%d/%Y")
date_cut = last_day_of_previous_month.strftime("%m.%y")

for co_code in co_codes:
    # change the path to relative
    CSV_file_name = "C:/Users/zawadzmi/Desktop/WD MEC 08.17/Script/Result/to_upload/{}_Accrual_WD_{}.csv".format(co_code, date_cut)
    JE_template_sheet_name = "Document Template"

    # here put the code for filling out the template

    #JE_csv["Account"] = something
    # ...
    # fill out each required field in the CSV DF using the groupby
#print(pivot.unstack())
#print(pivot1.head(3))