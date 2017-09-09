import pandas as pd
import openpyxl as oxl


def generate_csv(cc):  # cc = Company Code

    CSV_file_name = "../Script/Result/to_upload/{}_Accrual_WD_{}.csv".format(cc, date_cut)
    # generate the form
    JE_csv = pd.DataFrame(columns=JE_csv_columns)
    cur_cc_data = grouped_by_cc.get_group(cc)
    grouped_by_checksum = cur_cc_data.groupby(["Checksum"])
    posting_month = last_day_of_previous_month.strftime("%b")
    posting_year = last_day_of_previous_month.strftime("%Y")

    start_row = 0
    for k, g in grouped_by_checksum:
        #print(k)
        #print(g)
        for i in range(start_row, start_row + len(g)):
            # for each line for a given checksum (BA and PC combination), retrieve its Acc# culumn value and input it
            # into the next free cell in the "ACCOUNT" column in the JE csv form
            JE_csv.loc[i, "ACCOUNT"] = g.iloc[i - start_row]["Acc#"]
            JE_csv.loc[i, "DEBIT"] = g.iloc[i - start_row]["Net Amount LC"]
            JE_csv.loc[i, "LINE MEMO"] = g.iloc[i - start_row]["Expense Report Number"] + " Accrual"
            JE_csv.loc[i, "BUSINESS AREA"] = g.iloc[i - start_row]["Checksum"][-4:]  # BA is the last 4 chars of checksum
            JE_csv.loc[i, "PROFIT CENTER"] = g.iloc[i - start_row]["Checksum"][:5]  # PC is the first 5 chars of checksum

            # make sure this is correctly read by Excel
            JE_csv.loc[i, "DATE"] = last_day_of_previous_month

            JE_csv.loc[i, "SUBSIDIARY"] = cc
            JE_csv.loc[i, "MEMO"] = "WD {} ACCRUALS {} FY{}".format(cc, posting_month, posting_year)
        start_row += len(g)
    if cc == "BESA":
        print(JE_csv)

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