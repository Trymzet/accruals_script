import pandas as pd
import openpyxl as oxl
import numpy as np
from urllib.request import urlopen
from json import loads

# TODO: delete below line
WD_report_name = "EXP031-RPT-Process-Accruals_with_Expense_Report.xlsx"


def generate_exchange_rates():
    # https://openexchangerates.org
    exchange_rates_api_key = "11f20df062814531be891cc0173702a6"
    api_call = f"https://openexchangerates.org/api/latest.json?app_id={exchange_rates_api_key}"
    rates_api_response = urlopen(api_call)
    rates_api_response_str = rates_api_response.read().decode("ascii")
    rates_api_response_dict = loads(rates_api_response_str)
    rates = rates_api_response_dict["rates"]

    # feel free to update company codes/currencies
    currencies_in_scope = {"AUSA": "AUD", "BESA": "EUR", "BGSA": "BGN", "BRSA": "BRL", "CASA": "CAD", "CHSD": "CHF", "CNSA": "CNY",
                           "CRSB": "CRC", "CZSA": "CZK", "DESA": "EUR", "DKSA": "DKK", "ESSA": "EUR", "FRSA": "EUR", "GBF0": "USD",
                           "GBSA": "GBP", "IESA": "EUR", "IESB": "EUR", "ILSA": "ILS", "ILSB": "ILS", "INSA": "INR", "INSB": "INR",
                           "INSD": "INR", "ITSA": "EUR", "JPSA": "JPY", "LUSB": "EUR", "MXSC": "MXN", "NLSC": "EUR", "PHSB": "PHP",
                           "PLSA": "PLN", "PRSA": "PYG", "ROSA": "RON", "RUSA": "RUB", "SESA": "SEK", "TRSA": "TRY", "USMS": "USD",
                           "USSM": "USD", "USSN": "USD"}

    exchange_rates_to_usd = {}
    for company_code in currencies_in_scope:
        currency = currencies_in_scope[company_code]
        # the rates from API are from USD to x; we need from x to USD
        try:
            exchange_rate_to_usd = 1/rates[currency]
        except:
            continue
        if company_code in exchange_rates_to_usd:
            continue
        else:
            exchange_rates_to_usd[company_code] = exchange_rate_to_usd
    return exchange_rates_to_usd



def generate_csv(cc):  # cc = Company Code

    CSV_file_name = "../Script/Result/to_upload/{}_Accrual_WD_{}.csv".format(cc, date_cut)
    JE_csv = pd.DataFrame(columns=JE_csv_columns)  # TODO make one template DF and use it for all ccs
    cur_cc_data = grouped_by_cc.get_group(cc)
    grouped_by_checksum = cur_cc_data.groupby(["Checksum"])
    posting_month = last_day_of_previous_month.strftime("%b")
    posting_year = last_day_of_previous_month.strftime("%Y")
    posting_period = "{} {}".format(posting_month, posting_year)
    # this is a way to track row number, so that groups can be input to consecutive rows
    cur_group_start_row = 0

    for checksum, g in grouped_by_checksum:
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

        # TODO: final result should be grouped by sum per report number
        # cur_checksum_data = g
        # grouped_by_expense_report = cur_checksum_data.groupby(["Expense Report Number"])
        # report_sum_by_expense_report = grouped_by_expense_report.sum()

    JE_amount_local = JE_csv["CREDIT"].sum(skipna=True)
    exchange_rates = generate_exchange_rates()
    amount_in_usd =  JE_amount_local * exchange_rates[cc]
    to_generate = []

    # company requirement
    if amount_in_usd > 5000:
        to_generate.append(cc)

    if cc in to_generate:
        JE_csv.to_csv(CSV_file_name, index=False)
        print("{} CSV file generated :)".format(cc))

preprocessed_WD_report = pd.read_excel("../Script/Result/{}".format(WD_report_name))
WD_report_groupby_input = preprocessed_WD_report[["Entity Code", "Checksum", "Acc#", "Expense Report Number", "Net Amount LC", "MRU", "Functional Area"]]
grouped_by_cc = WD_report_groupby_input.groupby("Entity Code", as_index=False)

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
