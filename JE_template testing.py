import pandas as pd
import openpyxl as oxl

WD_report_name = "EXP031-RPT-Process-Accruals_with_Expense_Report.xlsx"
master_file_name = "WD_Accruals_Master.xlsx"
AP_account = 25702400  # the account from which the money will flow


last_day_of_previous_month = (pd.to_datetime("today") - pd.tseries.offsets.MonthEnd(1))#.strftime("%m/%d/%Y")
print(last_day_of_previous_month)
excel_format = last_day_of_previous_month - pd.to_datetime("18991231")
print(int(str(excel_format)[:5]))
print(pd.Series(excel_format).dt.days)

JE_template_sheet_name = "Document Template"
#master_file = oxl.load_workbook(master_file_name, keep_vba=True)
#JE_template = master_file[JE_template_sheet_name]

JE_csv_columns = ["ACCOUNT", "DEBIT", "CREDIT", "TAX CODE", "LINE MEMO", "MRU", "BUSINESS AREA", "PROFIT CENTER", "FUNCTIONAL AREA",
                  "DATE", "ACCOUNTING BOOK", "SUBSIDIARY", "CURRENCY", "MEMO", "REVERSAL DATE", "TO SUBSIDIARY", "UNIQUE ID"]

JE_csv = pd.DataFrame(columns=JE_csv_columns)
