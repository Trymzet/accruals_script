### Author: Michal Zawadzki, michalmzawadzki@gmail.com                  ###
### Run the script from the same directory as WD report and master file ###

import pandas as pd
import openpyxl
from subprocess import call
from os import remove

WD_report_name = "EXP031-RPT-Process-Accruals_with_Expense_Report.xlsx"
WD2_report_name = "EXP032-RPT-Process-Accruals-_No_Expense.xlsx"
AP_account = 25702400  # the account from which the money will flow
generic_GL_account = 46540000


def generate_vbs_script():
    vbscript = """if WScript.Arguments.Count < 3 Then
        WScript.Echo "Please specify the source and the destination files. Usage: ExcelToCsv <xls/xlsx source file> <csv destination file> <worksheet number (starts at 1)>"
        Wscript.Quit
    End If

    csv_format = 6

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    src_file = objFSO.GetAbsolutePathName(Wscript.Arguments.Item(0))
    dest_file = objFSO.GetAbsolutePathName(WScript.Arguments.Item(1))
    worksheet_number = CInt(WScript.Arguments.Item(2))

    Dim oExcel
    Set oExcel = CreateObject("Excel.Application")

    Dim oBook   
    Set oBook = oExcel.Workbooks.Open(src_file)
    oBook.Worksheets(worksheet_number).Activate

    oBook.SaveAs dest_file, csv_format

    oBook.Close False
    oExcel.Quit
    """

    with open("ExcelToCsv.vbs", "wb") as f:
        f.write(vbscript.encode("utf-8"))


def load_csv(excel_file_name, has_sheets=False, skiprows=None):
    if has_sheets:
        # sheet numbers to use; using the first three, hence the fixed numbers; fifth sheet in Master file is the template,
        # which has to be parsed as xlsx
        sheets = map(str, range(1, 4))
        sheet_dataframes = []
        for sheet in sheets:
            csv_file_name = "../Script/{}{}".format(sheet, ".csv")
            call(["cscript.exe", "../Script/ExcelToCsv.vbs", excel_file_name, csv_file_name, sheet, r"//B"])
            sheet_dataframe = pd.read_csv(csv_file_name, encoding="latin-1")
            sheet_dataframes.append(sheet_dataframe)
        return tuple(sheet_dataframes)
    else:
        csv_file_name = excel_file_name[:-4] + "csv"
        # //B is for batch mode; this is to avoid spam on the console
        call(["cscript.exe", "../Script/ExcelToCsv.vbs", excel_file_name, csv_file_name, str(1), r"//B"])
        if skiprows:
            try:
                data = pd.read_csv(csv_file_name, skiprows=skiprows, encoding="latin-1")
            except:
                print("Something went wrong... make sure report names weren't changed or debug the load_csv function")
        else:
            try:
                data = pd.read_csv(csv_file_name, encoding="latin-1")
            except:
                print("Something went wrong... make sure report names weren't changed or debug the load_csv function")
        return data


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
    else:
        try:
            wb = pd.read_excel(file_name, na_values="")
        except:
            print("File {} not found in current directory.".format(file_name))
    return wb


def load_all():
    master_file_name = "WD_Accruals_Master.xlsm"
    file_names = [WD_report_name, WD2_report_name, master_file_name]
    dataframes = []

    # this will be used by each load_csv() call to convert every .xlsx to .csv for faster loading
    generate_vbs_script()

    for file_name in file_names:
        if file_name == WD_report_name:
            df = load_csv(file_name, skiprows=[0])
        elif file_name == WD2_report_name:
            df = load_csv(file_name)
        else:
            sheet1, sheet2, sheet3 = load_csv(file_name, has_sheets=True)
            dataframes.extend([sheet1, sheet2, sheet3])
            return dataframes
        dataframes.append(df)
    return dataframes


def initial_cleanup():
    global WD_report, WD2_report
    # no longer needed
    remove("ExcelToCsv.vbs")
    # remove rows with total amount 0 or less
    WD_report = WD_report[WD_report["Net Amount LC"] > 0]
    WD2_report = WD2_report[WD2_report["Billing Amount"] > 0]
    # delete any rows that don't have a cost center specified
    WD_report.dropna(subset=["Cost Center"], inplace=True)
    WD2_report.dropna(subset=["Report Cost Location"], inplace=True)
    # delete the duplicate cost centers/descriptions inside Cost Center column
    WD_report["Cost Center"] = WD_report["Cost Center"].astype("str").map(lambda x: x.split()[0])
    WD2_report["Report Cost Location"] = WD2_report["Report Cost Location"].astype("str").map(lambda x: x.split()[0])


def vlookup(what, left_on, right_on):
    result = WD_report.merge(what, left_on=left_on, right_on=right_on, how="left")
    return result


def run_vlookups():
    global WD_report, WD2_report
    accounts = pd.DataFrame(accounts_file["Acc#"]).astype(int)
    ba_pc_to_join = ba_pc_file[["Business Area", "HPE Profit Center"]]

    WD_report = vlookup(accounts, WD_report["Expense Item"], accounts_file["Expense Item name"])
    # the account number is provided separately for each country. However, all countries have the same account for a given category, so we need to remove these duplicate rows.
    # in case any country has a separate account for a given category in the future, the script will still work
    WD_report.drop_duplicates(inplace=True)
    WD_report = vlookup(ba_pc_to_join, WD_report["Cost Center"], ba_pc_file["Legacy Cost Center"])
    WD2_report = vlookup(ba_pc_to_join, WD2_report["Report Cost Location"], ba_pc_file["Legacy Cost Center"])
    print(WD2_report.head())


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
    WD_report["Acc#"] = WD_report["Acc#"].map(int)
    # the accounts vlookup finds all possible accounts for a category, so this also needs to be filtered by subsidiary
    duplicates = WD_report.duplicated(subset=["Expense Report", "Expense Item", "Net Amount LC"])
    print(duplicates[duplicates == True].shape[0])
    duplicates = WD_report[duplicates == True]
    print(duplicates)

    duplicates.to_csv("duplicates.xlsx")


WD_report, WD2_report, ba_pc_file, accounts_file, JE_template_file = load_all()
print("All files loaded :)")
initial_cleanup()
print("Data cleaned :)")
run_vlookups()
final_cleanup()



compare = pd.read_excel("compare.xlsx")
compare["Acc#"] = compare["Acc#"].map(int)
d = {"WD_report": WD_report.astype(str), "compare": compare.astype(str)}
df = pd.concat(d)
df.drop_duplicates(keep=False, inplace=True)


# TODO: if the account / profit center / ba is not found, add those lines to a "not found" file for a manual check


WD_report["Checksum"] = WD_report["HPE Profit Center"].astype(str) + WD_report["Business Area"]

temp = "C:/Users/zawadzmi/Desktop/WD MEC 08.17/Script/Result/" + WD_report_name

with pd.ExcelWriter(temp) as writer:
    WD_report.to_excel(writer, engine="openpyxl", index=False)

with pd.ExcelWriter("concatenated.xlsx") as writer:
    df.to_excel(writer, engine="openpyxl", index=False)