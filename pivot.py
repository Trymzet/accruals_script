import pandas as pd
import openpyxl as oxl


def generate_csv(cc):  # cc = Company Code

    CSV_file_name = "../Script/Result/to_upload/{}_Accrual_WD_{}.csv".format(cc, date_cut)
    JE_csv = pd.DataFrame(columns=JE_csv_columns)  # TODO make one template DF and use it for all ccs
    cur_cc_data = grouped_by_cc.get_group(cc)
    grouped_by_checksum = cur_cc_data.groupby(["Checksum"])
    posting_month = last_day_of_previous_month.strftime("%b")
    posting_year = last_day_of_previous_month.strftime("%Y")
    posting_period = "{} {}".format(posting_month, posting_year)


# TODO: NEED TO ADD ADDITIONAL GRUPBY - FA or MRU?
    cur_group_start_row = 0

    for checksum, g in grouped_by_checksum:
        print(g)
        business_area = checksum[-4:]  # BA is the last 4 chars of checksum
        profit_center = checksum[:5]  # PC is the first 5 chars of checksum
        general_description = "WD {} ACCRUALS {} FY{}".format(cc, posting_month, posting_year)
        for i in range(cur_group_start_row, cur_group_start_row + len(g)):
            # for each line for a given checksum (BA and PC combination), retrieve its Acc# culumn value and input it
            # into the next free cell in the "ACCOUNT" column in the JE csv form
            JE_csv.loc[i, "ACCOUNT"] = g.iloc[i - cur_group_start_row]["Acc#"]
            JE_csv.loc[i, "DEBIT"] = g.iloc[i - cur_group_start_row]["Net Amount LC"]
            JE_csv.loc[i, "LINE MEMO"] = g.iloc[i - cur_group_start_row]["Expense Report Number"] + " Accrual"
            # Note that even though the template has a TRANSACTION DATE - DAY field, it still passes the whole date in mm/dd/YYYY format
            JE_csv.loc[i, "DATE"] = last_day_of_previous_month.strftime("%m/%d/%Y")
            JE_csv.loc[i, "POSTING PERIOD"] = posting_period
            JE_csv.loc[i, "SUBSIDIARY"] = cc
            JE_csv.loc[i, "MEMO"] = general_description
            JE_csv.loc[i, "REVERSAL DATE"] = first_day_of_current_month
            JE_csv.loc[i, "MRU"] = g.iloc[i - cur_group_start_row]["MRU"]
            JE_csv.loc[i, "FUNCTIONAL AREA"] = g.iloc[i - cur_group_start_row]["Functional Area"]

        # TODO: calculate current JE's total in USD by using OANDA's API -- write get_oanda_exchange_rate(cc)
        # cur_ccs_exchange_rate = get_oanda_exchange_rate(cc)
        # JE_csv_sum_USD = JE_csv_sum * cur_ccs_exchange_rate
        # if JE_csv_sum_USD < 5000:
        #    print("Total amount of accruals for {} is lower than 5000 USD, hence not generating JE".format(cc))
        #    return

        # here we're filling out the AP account row
        last_group_start_row = cur_group_start_row
        cur_group_start_row += len(g)
        JE_csv.loc[cur_group_start_row, "ACCOUNT"] = 25702400
        JE_csv.loc[cur_group_start_row, "CREDIT"] = JE_csv.loc[last_group_start_row:cur_group_start_row, "DEBIT"].sum()
        JE_csv.loc[cur_group_start_row, "LINE MEMO"] = general_description
        JE_csv.loc[cur_group_start_row, "BUSINESS AREA"] = business_area
        JE_csv.loc[cur_group_start_row, "PROFIT CENTER"] = profit_center
        JE_csv.loc[cur_group_start_row, "DATE"] = last_day_of_previous_month.strftime("%m/%d/%Y")
        JE_csv.loc[cur_group_start_row, "POSTING PERIOD"] = posting_period
        JE_csv.loc[cur_group_start_row, "SUBSIDIARY"] = cc
        JE_csv.loc[cur_group_start_row, "MEMO"] = general_description
        JE_csv.loc[cur_group_start_row, "REVERSAL DATE"] = first_day_of_current_month
        cur_group_start_row += 1

        # TODO: final result should be grouped by sum per report no, not per account -> delete below code
        cur_checksum_data = g
        grouped_by_acc = cur_checksum_data.groupby(["Expense Report Number", "Acc#"])
        report_sum_by_acc = grouped_by_acc.sum()

        #print(report_sum_by_acc)

        # use merge to vlookup final sum per account number -
        # 1. Generate a DataFrame from report_sum_by_acc
        # 2. JE_csv.merge(report_sum_by_acc["Net Amount LC"], left_on="Expense Report Number", right_on= "Expense Report Number"

        #for kee, grup in grouped_by_acc:
            #print(kee)
            #print(grup)

    JE_csv.to_csv(CSV_file_name, index=False)
    print("{} CSV file generated :)".format(cc))


# clean = pd.read_excel("clean.xlsx") tested, works
clean = pd.read_excel("C:/Users/zawadzmi/Desktop/WD MEC 08.17/Script/Result/" + "EXP031-RPT-Process-Accruals_with_Expense_Report.xlsx")
template_data = clean[["Entity Code", "Checksum", "Acc#", "Expense Report Number", "Net Amount LC", "MRU", "Functional Area"]]
grouped_by_cc = template_data.groupby("Entity Code", as_index=False)

JE_csv_columns = ["ACCOUNT", "DEBIT", "CREDIT", "TAX CODE", "LINE MEMO", "MRU", "BUSINESS AREA", "PROFIT CENTER", "FUNCTIONAL AREA",
                  "DATE", "POSTING PERIOD", "ACCOUNTING BOOK", "SUBSIDIARY", "CURRENCY", "MEMO", "REVERSAL DATE", "TO SUBSIDIARY",
                  "TRADING PARTNER", "TRADING PARTNER CODE", "UNIQUE ID"]
last_day_of_previous_month = pd.to_datetime("today") - pd.tseries.offsets.MonthEnd(1)
date_cut = last_day_of_previous_month.strftime("%m.%y")
first_day_of_current_month = pd.to_datetime("today").replace(day=1).strftime("%m/%d/%Y")
JE_template_sheet_name = "Document Template"

for key, group in grouped_by_cc:
    company_code = key
    generate_csv(company_code)



