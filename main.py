### Author: Michal Zawadzki, michalmzawadzki@gmail.com                  ###
### Run the script from the same directory as WD report and master file ###

import pandas as pd
import openpyxl

WD_report_name = "EXP031-RPT-Process-Accruals_with_Expense_Report.xlsx"
AP_account = 25702400  # the account from which the money will flow


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


def load_all():
    master_file_name = "WD_Accruals_Master.xlsx"
    accounts_sheet_name = "GL_accounts_by_category"
    ba_pc_sheet_name = "CC_to_BA_PC"
    JE_template_sheet_name = "FAST JE Template"

    file_names = [WD_report_name, accounts_sheet_name, ba_pc_sheet_name, JE_template_sheet_name]
    dataframes = []

    for file_name in file_names:
        if file_name == WD_report_name:
            df = load_workbook(file_name, skiprows=[0])
            dataframes.append(df)
        else:
            df = load_workbook(master_file_name, sheetname=file_name)
            dataframes.append(df)
    return dataframes


def initial_cleanup():
    global WD_report
    # remove rows with total amount 0 or less
    WD_report = WD_report[WD_report["Net Amount LC"] > 0]
    # delete any rows that don't have a cost center specified
    WD_report.dropna(subset=["Cost Center"], inplace=True)
    # delete the duplicate cost centers/descriptions inside Cost Center column
    WD_report["Cost Center"] = WD_report["Cost Center"].astype("str").map(lambda x: x.split()[0])


def vlookup(what, left_on, right_on):
    result = WD_report.merge(what, left_on=left_on , right_on=right_on, how="left")
    return result


def run_vlookups():
    global WD_report
    accounts = pd.DataFrame(accounts_file["Acc#"]).astype(int)
    ba_pc_to_join = ba_pc_file[["Business Area", "HPE Profit Center"]]

    WD_report = vlookup(accounts, WD_report["Expense Item"], accounts_file["Expense Item name"])
    # the account number is provided separately for each country. However, all countries have the same account for a given category, so we need to remove these duplicate rows.
    # in case any country has a separate account for a given category in the future, the script will still work
    WD_report.drop_duplicates(inplace=True)


    print(WD_report.shape[0], "After dropping duplicate accounts")

    # TODO: check all the columns UP TO Acc# for duplicates - if they only differ in Acc#, it means the acc was different for the same category!
    # TODO: there should be 22 such situations

    WD_report = vlookup(ba_pc_to_join, WD_report["Cost Center"], ba_pc_file["Legacy Cost Center"])


def final_cleanup():
    global WD_report
    travel_journal_item_account = 46540000
    company_celebration_account = 46900000
    german_cost_account = 46920000

    # add vlookup exceptions
    no_of_items = WD_report.shape[0]
    for row_index in range(no_of_items):
        category = str(WD_report["Expense Item"].iloc[row_index])  # for some reason this column is loaded as float, hence the str()
        if "Travel Journal Item" in category:
            WD_report.loc[row_index, "Acc#"] = travel_journal_item_account
            # WD_report.set_value(index, "Acc#", travel_journal_item_account)
        if "Company Celebration" in category:
            # WD_report.set_value(index, "Acc#", company_celebration_account)
            WD_report.loc[row_index, "Acc#"] = company_celebration_account

    # this is to stop Excel from reding e.g. 2E00 as a number in scientific notation
    WD_report["Business Area"] = WD_report["Business Area"].map(str)

    # note that this also overrides the above two exceptions, which are changed to the german account
    WD_report.loc[WD_report["Entity Code"] == "DESA", "Acc#"] = german_cost_account
    # ensure account number is an integer


WD_report, accounts_file, ba_pc_file, JE_template = load_all()
print("All files loaded :)")
initial_cleanup()
print("Data cleaned :)")
run_vlookups()
final_cleanup()


# TODO: if the account / profit center / ba is not found, add those lines to a "not found" file for a manual check


WD_report["Checksum"] = WD_report["HPE Profit Center"].astype(str) + WD_report["Business Area"]

temp = "C:/Users/zawadzmi/Desktop/WD MEC 08.17/Script/Result/" + WD_report_name

with pd.ExcelWriter(temp) as writer:
    WD_report.to_excel(writer, engine="openpyxl", index=False)

WD_report.to_csv("Prepared.csv")