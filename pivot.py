import pandas as pd
import openpyxl as oxl


def generate_csv(cc):  # cc = Company Code

    CSV_file_name = "../Script/Result/to_upload/{}_Accrual_WD_{}.csv".format(cc, date_cut)
    JE_csv = pd.DataFrame(columns=JE_csv_columns)  # TODO make one template DF and use it for all ccs
    cur_cc_data = grouped_by_cc.get_group(cc)
    grouped_by_checksum = cur_cc_data.groupby(["Checksum"])
    posting_month = last_day_of_previous_month.strftime("%b")
    posting_year = last_day_of_previous_month.strftime("%Y")

    start_row = 0
    for checksum, g in grouped_by_checksum:
        business_area = checksum[-4:]  # BA is the last 4 chars of checksum
        profit_center = checksum[:5]  # PC is the first 5 chars of checksum
        general_description = "WD {} ACCRUALS {} FY{}".format(cc, posting_month, posting_year)
        for i in range(start_row, start_row + len(g)):
            # for each line for a given checksum (BA and PC combination), retrieve its Acc# culumn value and input it
            # into the next free cell in the "ACCOUNT" column in the JE csv form
            JE_csv.loc[i, "ACCOUNT"] = g.iloc[i - start_row]["Acc#"]
            JE_csv.loc[i, "DEBIT"] = g.iloc[i - start_row]["Net Amount LC"]
            JE_csv.loc[i, "LINE MEMO"] = g.iloc[i - start_row]["Expense Report Number"] + " Accrual"
            JE_csv.loc[i, "DATE"] = last_day_of_previous_month.strftime("%m/%d/%Y")
            JE_csv.loc[i, "SUBSIDIARY"] = cc
            JE_csv.loc[i, "MEMO"] = general_description
            JE_csv.loc[i, "REVERSAL DATE"] = first_day_of_current_month

        # here we're filling out the AP account row
        a = start_row  # use a proper variable name
        start_row += len(g)
        JE_csv.loc[start_row, "ACCOUNT"] = 25702400
        JE_csv.loc[start_row, "CREDIT"] = JE_csv.loc[a:start_row, "DEBIT"].sum()
        JE_csv.loc[start_row, "LINE MEMO"] = general_description
        JE_csv.loc[start_row, "BUSINESS AREA"] = business_area
        JE_csv.loc[start_row, "PROFIT CENTER"] = profit_center
        JE_csv.loc[start_row, "DATE"] = last_day_of_previous_month.strftime("%m/%d/%Y")
        JE_csv.loc[start_row, "SUBSIDIARY"] = cc
        JE_csv.loc[start_row, "MEMO"] = general_description
        JE_csv.loc[start_row, "REVERSAL DATE"] = first_day_of_current_month
        start_row += 1
        print("{} CSV file generated :)".format(cc))

    JE_csv.to_csv(CSV_file_name, index=False)


clean = pd.read_excel("clean.xlsx")
template_data = clean[["Entity Code", "Checksum", "Acc#", "Expense Report Number", "Net Amount LC"]]
grouped_by_cc = template_data.groupby("Entity Code", as_index=False)

JE_csv_columns = ["ACCOUNT", "DEBIT", "CREDIT", "TAX CODE", "LINE MEMO", "MRU", "BUSINESS AREA", "PROFIT CENTER", "FUNCTIONAL AREA",
                  "DATE", "ACCOUNTING BOOK", "SUBSIDIARY", "CURRENCY", "MEMO", "REVERSAL DATE", "TO SUBSIDIARY", "UNIQUE ID"]
last_day_of_previous_month = pd.to_datetime("today") - pd.tseries.offsets.MonthEnd(1)
date_cut = last_day_of_previous_month.strftime("%m.%y")
first_day_of_current_month = pd.to_datetime("today").replace(day=1).strftime("%m/%d/%Y")
JE_template_sheet_name = "Document Template"

for key, group in grouped_by_cc:
    company_code = key
    generate_csv(company_code)
