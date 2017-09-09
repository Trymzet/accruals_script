import pandas as pd
import openpyxl as oxl


def generate_csv(cc): # cc = Company Code

    CSV_file_name = "../Script/Result/to_upload/{}_Accrual_WD_{}.csv".format(cc, date_cut)
    # generate the form
    JE_csv = pd.DataFrame(columns=JE_csv_columns)

    cur_cc_data = grouped_by_cc.get_group(cc)
    print("AAAAAAAAAAAA")
    #print(cur_cc_data)
    grouped_by_checksum = cur_cc_data.groupby(["Checksum", "Expense Report Number"])
    for k, g in grouped_by_checksum:
        print(g)

    # here put the code for filling out the template

    #JE_csv.to_csv(CSV_file_name, index=False)


clean = pd.read_excel("clean.xlsx")
cleaner = clean[["Entity Code", "Checksum", "Acc#", "Expense Report Number", "Net Amount LC"]]
grouped_by_cc = cleaner.groupby("Entity Code", as_index=False)

ccs = []
for key, group in grouped_by_cc:
    cc = key
    if cc not in ccs:
        ccs.append(cc)


JE_csv_columns = ["ACCOUNT", "DEBIT", "CREDIT", "TAX CODE", "LINE MEMO", "MRU", "BUSINESS AREA", "PROFIT CENTER", "FUNCTIONAL AREA",
                  "DATE", "ACCOUNTING BOOK", "SUBSIDIARY", "CURRENCY", "MEMO", "REVERSAL DATE", "TO SUBSIDIARY", "UNIQUE ID"]
last_day_of_previous_month = (pd.to_datetime("today") - pd.tseries.offsets.MonthEnd(1))#.strftime("%m/%d/%Y")
date_cut = last_day_of_previous_month.strftime("%m.%y")
JE_template_sheet_name = "Document Template"


for cc in ccs:
    generate_csv(cc)