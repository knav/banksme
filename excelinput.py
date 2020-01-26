import pandas as pd
import numpy as np
import sys
from company import Company
from openpyxl import Workbook

# Path to the file to be analysed
# wb_path = r'/Users/navneeeth99/Desktop/Financial Reports/DBS 1.xls'
wb_path = r'C:\Users\Navneeth\Documents\WORK\BankSME\Programming\Financial Reports\DBS 1.xls'

# Find the worksheets with the balance sheet and income statement.
# If unable to find, get user input.
wb = pd.ExcelFile(wb_path)
wb_bs = ""
wb_is = ""
wb_bs_tags = ['bal', 'sh']
wb_is_tags = ['inc', 'st']
alt_wb_is_tags = ['p&l']
for name in wb.sheet_names:
    if all(i in name.casefold() for i in wb_bs_tags):
        wb_bs = name
    elif all(i in name.casefold() for i in wb_is_tags):
        wb_is = name
    elif any(i in name.casefold() for i in alt_wb_is_tags):
        wb_is = name

if wb_bs == "":
    while wb_bs == "":
        wb_bs = input("Unable to determine which tab contains balance sheet. "
                      "Please state the name of the relevant sheet: ")
        if wb_bs not in wb.sheet_names:
            print("Invalid entry.")
            continue

if wb_is == "":
    while wb_is == "":
        wb_is = input("Unable to determine which tab contains income statement. "
                      "Please state the name of the relevant sheet: ")
        if wb_is not in wb.sheet_names:
            print("Invalid entry.")
            continue

# Read and do basic cleanup of excel file
bsheet = pd.read_excel(wb_path, sheet_name=wb_bs, header=None)
bsheet.replace(['NaN'], np.nan, inplace=True)                                   # Replace blanks with np.nan for dropna
bsheet.dropna(axis=0, how='all', inplace=True)                                  # Drop fully blank rows
bsheet.dropna(axis=1, how='any', thresh=0.5*bsheet.shape[0], inplace=True)      # Drop columns that are >half-empty
bsheet.columns = range(bsheet.shape[1])                                         # Renumber the remaining columns

# If the data columns in a row are all blank, delete the row
na_subset = range(1, bsheet.shape[1])
bsheet.dropna(axis=0, how='all', subset=na_subset, inplace=True)

# Renumber the remaining rows
bsheet.reset_index(drop=True, inplace=True)

# If there is more than one column of data, ask for user to input which column of data to use
col_names = bsheet.iloc[0, 1:].to_list()
col_names = [str(i) for i in col_names]
cnt = 1
for j in col_names:
    print(str(cnt) + " '" + j + "'")
    cnt += 1

if len(bsheet.columns[1:].to_list()) > 1:
    cfmed = False
    while not cfmed:
        keep_index = int(input("From the list of columns above, enter the index of the column to be used: "))
        if 0 < keep_index <= len(col_names):
            bsheet = bsheet.iloc[:, [0, keep_index]]
            col_names = []
            cfmed = True
        else:
            print("Invalid entry")
            print()
            continue

# Delete any rows where the label column is empty
bsheet = bsheet[bsheet[0].notnull()]

# Get list of all the labels in the first column
label_col = bsheet.iloc[:, 0]
label_list = label_col.tolist()

# Set up key dictionaries --> from balance sheet
assets = {'Current Assets': 0,
          'Accounts Receivable': 0,
          'Marketable Securities': 0,
          'Loans': 0,
          'Cash': 0,
          'Inventory': 0,
          'PPE': 0}

liabeq = {'Current Liabilities': 0,
          'Long-term Debt': 0,
          'Deposits': 0,
          'Total Equity': 0,
          'Total Liabilities': 0}

# Labels to search for each value
alt_assets = {'Current Assets': ["Total Current Assets", "Net Current Assets", "Current Assets"],
              'Accounts Receivable': ["Accounts Receivable"],
              'Marketable Securities': ["Marketable Securities", "Short Term Investments", "Short-term Investments"],
              'Loans': ["Loans"],
              'Cash': ["Cash and Equivalents", "Cash and cash equivalents", "Cash"],
              'Inventory': ["Inventory", "Inventories"],
              'PPE': ["Property, Plant & Equipment", "PPE", "Property and equipment"],
              'Total Assets': ["Total Assets", "Net Assets"]}

alt_liabeq = {'Current Liabilities': ["Total Current Liabilities", "Net Current Liabilities", "Current Liabilities"],
              'Long-term Debt': ["Long Term Debt", "Long-term Debt"],
              'Deposits': ["Deposits"],
              'Total Equity': ["Total Equity", "Net Equity", "Shareholders' Equity", "Total Shareholders' Equity",
                               "Net Shareholders' Equity", "Total stockholders' equity", "Net stockholders' equity"],
              'Total Liabilities': ["Total Liabilities"],
              'Total Liabilities and Equity': ["Total Liabilities And Equity", "Net Liabilities And Equity",
                                               "Total Liabilities And Shareholders' Equity",
                                               "Net Liabilities And Shareholders' Equity",
                                               "Total Liabilities And Stockholders' Equity",
                                               "Net Liabilities And Stockholders' Equity"]}

# Instantiate total assets and total liabilities and equities figure separately - IMPT BASE NUMBERS
total_assets = 0
total_liabeq = 0

# Establish value of total assets (used for calc of %)
totasset_labels = [s for s in label_list if any(xs.casefold() in s.casefold() for xs in alt_assets['Total Assets'])]
totasset_indexes = [label_list.index(n) for n in totasset_labels]

if len(totasset_indexes) > 0:
    all_vals = []
    for v in totasset_indexes:
        all_vals.append(bsheet.iat[v, 1])

    total_assets = max(all_vals)
else:
    total_assets = input("Unable to find total assets figure. Please enter the value is $ millions: ")


# Establish value of total liabilities and equity (used for calc of %)
totliab_labels = [s for s in label_list if any(
    xs.casefold() in s.casefold() for xs in alt_liabeq['Total Liabilities and Equity'])]
totliab_indexes = [label_list.index(n) for n in totliab_labels]

if len(totliab_indexes) > 0:
    all_vals = []
    for v in totliab_indexes:
        all_vals.append(bsheet.iat[v, 1])

    total_liabeq = max(all_vals)
elif total_assets != 0:
    print("Unable to find total liabilities and equities figure. "
          "Proceeding with total assets figure.")
    total_liabeq = total_assets

# Safety check: total_assets = total_liabeq. Exit if mismatch
if total_assets != total_liabeq:
    print("Total assets and total liabilities and equity do not match!")
    print("Total assets: " + str(total_assets))
    print("Total liabilities and equity: " + str(total_liabeq))
    sys.exit()

# Loop through and fill up values, values that cannot be found are given np.nan
for a in assets:
    a_labels = [s for s in label_list if any(xs.casefold() in s.casefold() for xs in alt_assets[a])]
    a_indexes = [label_list.index(n) for n in a_labels]

    if len(a_indexes) > 0:
        all_vals = []
        for v in a_indexes:
            if bsheet.iat[v, 1] / total_assets < 1:
                all_vals.append(bsheet.iat[v, 1])

        # Note that besides total_assets/liabeq, all values are stored as percentages of total_assets/liabeq
        if all_vals != []:
            assets[a] = (max(all_vals)/total_assets)*100
        else:
            assets[a] = np.nan
    else:
        assets[a] = np.nan

assets['Total Assets'] = total_assets

for le in liabeq:
    le_labels = [s for s in label_list if any(xs.casefold() in s.casefold() for xs in alt_liabeq[le])]
    le_indexes = [label_list.index(n) for n in le_labels]

    if len(le_indexes) > 0:
        all_vals = []
        for v in le_indexes:
            if bsheet.iat[v, 1] / total_liabeq < 1:
                all_vals.append(bsheet.iat[v, 1])

        liabeq[le] = (max(all_vals)/total_liabeq)*100
    else:
        liabeq[le] = np.nan

liabeq['Total Liabilities and Equity'] = total_liabeq

###################### END OF BALANCE SHEET, START OF INCOME STATEMENT ######################

# Exact same extraction process as balance sheet
instat = pd.read_excel(wb_path, sheet_name=wb_is, header=None)
instat.replace(['NaN'], np.nan, inplace=True)
instat.dropna(axis=0, how='all', inplace=True)
instat.dropna(axis=1, how='any', thresh=0.5*instat.shape[0], inplace=True)
instat.columns = range(instat.shape[1])
na_subset = range(1, instat.shape[1])
instat.dropna(axis=0, how='all', subset=na_subset, inplace=True)

instat.reset_index(drop=True, inplace=True)

# Confirm column to use
col_names = instat.iloc[0, 1:].to_list()
col_names = [str(i) for i in col_names]
cnt = 1
for j in col_names:
    print(str(cnt) + " '" + j + "'")
    cnt += 1

if len(instat.columns[1:].to_list()) > 1:
    cfmed = False
    while not cfmed:
        keep_index = int(input("From the list of columns above, enter the index of the column to be used: "))
        if 0 < keep_index <= len(col_names):
            instat = instat.iloc[:, [0, keep_index]]
            col_names = []
            cfmed = True
        else:
            print("Invalid entry")
            print()
            continue

instat = instat[instat[0].notnull()]
instat_label_col = instat.iloc[:, 0]
instat_label_list = instat_label_col.tolist()

# Values needed from income statement
incst = {'Cost of Goods Sold': 0,
         'Net Profit': 0,
         'Revenue': 0,
         'Net interest income': 0,
         'Non-interest income': 0,
         'Interest Expense': 0,
         'EBIT': 0,
         'EBITDA': 0}

alt_incst = {'Cost of Goods Sold': ["Cost of Goods Sold", "Cost of Revenue", "Cost of Sales", "Cost of net revenues"],
             'Net Profit': ["Net Profit", "Total Profit", "Net Income", "Total Income"],
             'Revenue': ["Total revenue", "Net revenue", "Revenue"],
             'Net interest income':['Net interest income', 'Net-interest income'],
             'Non-interest income': ['Non-interest income', 'Noninterest income', 'Non interest income'],
             'Interest Expense': ["Interest Expense"],
             'EBIT': ["EBIT"],
             'EBITDA': ["EBITDA"]}

# Fill in all the values
for i in incst:
    i_labels = [s for s in instat_label_list if any(xs.casefold() in s.casefold() for xs in alt_incst[i])]
    i_indexes = [instat_label_list.index(n) for n in i_labels]

    if len(i_indexes) > 0:
        all_vals = []
        for v in i_indexes:
            all_vals.append(instat.iat[v, 1])

        incst[i] = max(all_vals)
    else:
        incst[i] = np.nan

# Check for early indicators of bank balance sheet and plug gaps for ROE calculation
if incst['Net interest income'] and incst['Non-interest income'] != np.nan and np.isnan(incst['Revenue']):
    incst['Revenue'] = incst['Net interest income'] + incst['Non-interest income']

print()
print(pd.DataFrame.from_dict(assets, orient="index", columns=['Values']))
print()
print(pd.DataFrame.from_dict(liabeq, orient="index", columns=['Values']))
print()
print(pd.DataFrame.from_dict(incst, orient="index", columns=['Values']))
print()

comp_data = Company(2010, assets, liabeq, incst)
comp_data.getindustry()

