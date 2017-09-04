import pandas as pd

def load_workbook(file_name, sheetname=None, skiprows=None):
    if skiprows:
        try:
            wb = pd.read_excel(file_name, skiprows=skiprows, na_values="")
        except FileNotFoundError:
            print("File {} not found in current directory.".format(file_name))

    if sheetname:
        try:
            wb = pd.read_excel(file_name, sheetname=sheetname, na_values="")
        except:
            print("Sheet {} has been renamed or deleted".format(sheetname))
    return wb


WD_report_name = "EXP031-RPT-Process-Accruals_with_Expense_Report.xlsx"
master_file_name = "WD_Accruals_Master.xlsx"
AP_account = 25702400  # the account from which the money will flow

JE_template_sheet_name = "FAST JE Template"
JE_template = load_workbook(master_file_name, sheetname=JE_template_sheet_name)

# fields to fill out in the template
posting_date_pos = JE_template.iloc[4][0]
subsidiary_pos = JE_template.iloc[5][0]
reversal_date_pos = JE_template.iloc[8][0]
header_text_pos = JE_template.iloc[9][0]
account_pos = JE_template.iloc[15][0]
debit_pos = JE_template.iloc[15][1]
credit_pos = JE_template.iloc[15][2]
memo_pos = JE_template.iloc[15][4]
ba_pos = JE_template.iloc[15][6]
pc_pos = JE_template.iloc[15][7]

print(posting_date_pos)
print(subsidiary_pos)
print(reversal_date_pos)
print(header_text_pos)
print(account_pos)
print(debit_pos)
print(credit_pos)
print(memo_pos)
print(ba_pos)
print(pc_pos)

JE_template.loc[15, 5] = "ALALAL"
JE_template.iloc[15][5] = "ACCRUAL"
print(JE_template.iloc[15][5])
#print(JE_template.loc[15][5:10])


