import pandas as pd
import numpy as np
import sys
from company import Company
from openpyxl import Workbook

# Path to the file to be analysed
# wb_path = r'/Users/navneeeth99/Desktop/Financial Reports/DBS 1.xls'
wb_path = r'C:\Users\Navneeth\Documents\WORK\BankSME\Programming\Financial Reports\Mcdonalds 1 - full.xls'

# Find the worksheets with the balance sheet and income statement.
# If unable to find, get user input.
wb = pd.ExcelFile(wb_path)
wb_bs = ""
wb_is = ""
wb_cf = ""
wb_bs_tags = ['bal', 'sh']
wb_is_tags = ['inc', 'st']
wb_cf_tags = ['ca', 'fl']
alt_wb_is_tags = ['p&l']
for name in wb.sheet_names:
    if all(i in name.casefold() for i in wb_bs_tags):
        wb_bs = name
    elif all(i in name.casefold() for i in wb_is_tags):
        wb_is = name
    elif all(i in name.casefold() for i in wb_cf_tags):
        wb_cf = name
    elif any(i in name.casefold() for i in alt_wb_is_tags):
        wb_is = name

if wb_bs == "":
    while wb_bs == "":
        wb_bs = input("Unable to determine which tab contains balance sheet. "
                      "Please state the name of the relevant sheet: ")
        if wb_bs not in wb.sheet_names:
            print("Invalid entry.")
            wb_bs = ""
            continue

if wb_is == "":
    while wb_is == "":
        wb_is = input("Unable to determine which tab contains income statement. "
                      "Please state the name of the relevant sheet: ")
        if wb_is not in wb.sheet_names:
            print("Invalid entry.")
            wb_is = ""
            continue

if wb_cf == "":
    while wb_cf == "":
        wb_cf = input("Unable to determine which tab contains cash flow statement. "
                      "Please state the name of the relevant sheet: ")
        if wb_cf not in wb.sheet_names:
            print("Invalid entry.")
            wb_cf = ""
            continue

# Read and do basic cleanup of excel file for balance sheet
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

# Find the first row where all the data columns are non-empty and set as column labels
reqindex = bsheet.dropna(axis=0, how='any', subset=na_subset).index[0]
new_col_names = bsheet.iloc[reqindex, 1:].tolist()
new_col_names.insert(0, "Labels")
print(new_col_names)
bsheet.columns = new_col_names
bsheet = bsheet.iloc[reqindex+1:, :]

# Delete any rows where the label column is empty
bsheet = bsheet[bsheet.iloc[:, 0].notnull()]

# Renumber the remaining rows
bsheet.index = range(0, bsheet.shape[0])

print()
print(bsheet)

# Read and do basic cleanup of excel file for income statement
instat = pd.read_excel(wb_path, sheet_name=wb_is, header=None)
instat.replace(['NaN'], np.nan, inplace=True)
instat.dropna(axis=0, how='all', inplace=True)
instat.dropna(axis=1, how='any', thresh=0.5*instat.shape[0], inplace=True)
instat.columns = range(instat.shape[1])

# If the data columns in a row are all blank, delete the row
na_subset = range(1, instat.shape[1])
instat.dropna(axis=0, how='all', subset=na_subset, inplace=True)

# Renumber the remaining rows
instat.reset_index(drop=True, inplace=True)

# Find the first row where all the data columns are non-empty and set as column labels
reqindex = instat.dropna(axis=0, how='any', subset=na_subset).index[0]
new_col_names = instat.iloc[reqindex, 1:].tolist()
new_col_names.insert(0, "Labels")
print(new_col_names)
instat.columns = new_col_names
instat = instat.iloc[reqindex+1:, :]

# Delete any rows where the label column is empty
instat = instat[instat.iloc[:, 0].notnull()]

# Renumber the remaining rows
instat.index = range(0, instat.shape[0])

print()
print(instat)

# Read and do basic cleanup of excel file for income statement
cfstat = pd.read_excel(wb_path, sheet_name=wb_is, header=None)
cfstat.replace(['NaN'], np.nan, inplace=True)
cfstat.dropna(axis=0, how='all', inplace=True)
cfstat.dropna(axis=1, how='any', thresh=0.5*cfstat.shape[0], inplace=True)
cfstat.columns = range(cfstat.shape[1])

# If the data columns in a row are all blank, delete the row
na_subset = range(1, cfstat.shape[1])
cfstat.dropna(axis=0, how='all', subset=na_subset, inplace=True)

# Renumber the remaining rows
cfstat.reset_index(drop=True, inplace=True)

# Find the first row where all the data columns are non-empty and set as column labels
reqindex = cfstat.dropna(axis=0, how='any', subset=na_subset).index[0]
new_col_names = cfstat.iloc[reqindex, 1:].tolist()
new_col_names.insert(0, "Labels")
print(new_col_names)
cfstat.columns = new_col_names
cfstat = cfstat.iloc[reqindex+1:, :]

# Delete any rows where the label column is empty
cfstat = cfstat[cfstat.iloc[:, 0].notnull()]

# Renumber the remaining rows
cfstat.index = range(0, cfstat.shape[0])

print()
print(cfstat)

# Merge the balance sheet and income statement info into one dataframe
mergedfs = pd.concat([bsheet, instat, cfstat], keys=["bs", "is", "cf"])
mergedfs.index = range(0, mergedfs.shape[0])

# Get list of all the labels in the first column
label_col = mergedfs.iloc[:, 0]
label_list = label_col.tolist()

# Setup the dictionaries for assets, liabeq and incst
assets = {'Current Assets': [], 'Accounts Receivable': [], 'Marketable Securities': [], 'Loans': [],
          'Cash': [], 'Inventory': [], 'PPE': [], 'Total Assets': []}

liabeq = {'Current Liabilities': [], 'Long-term Debt': [], 'Deposits': [], 'Total Equity': [],
                'Total Liabilities': [], 'Total Liabilities and Equity': []}

incst = {'Cost of Goods Sold': [], 'Net Profit': [], 'Revenue': [], 'Net interest income': [],
         'Non-interest income': [], 'Interest Expense': [], 'EBIT': [], 'EBITDA': []}

# Dictionary to save relevant index of columns
assets_index = {'Current Assets': 0,
                'Accounts Receivable': 0,
                'Marketable Securities': 0,
                'Loans': 0,
                'Cash': 0,
                'Inventory': 0,
                'PPE': 0,
                'Total Assets': 0}

liabeq_index = {'Current Liabilities': 0,
                'Long-term Debt': 0,
                'Deposits': 0,
                'Total Equity': 0,
                'Total Liabilities': 0,
                'Total Liabilities and Equity': 0}

incst_index = {'Cost of Goods Sold': 0,
               'Net Profit': 0,
               'Revenue': 0,
               'Net interest income': 0,
               'Non-interest income': 0,
               'Interest Expense': 0,
               'EBIT': 0,
               'EBITDA': 0}

# Labels to search for each value
alt_assets = {'Current Assets': ["Total Current Assets", "Current Assets"],
              'Accounts Receivable': ["Accounts Receivable"],
              'Marketable Securities': ["Marketable Securities", "Short Term Investments", "Short-term Investments"],
              'Loans': ["Loans"],
              'Cash': ["Cash and Equivalents", "Cash and cash equivalents", "Cash"],
              'Inventory': ["Inventory", "Inventories"],
              'PPE': ["Property, Plant & Equipment", "PPE", "Property and equipment"],
              'Total Assets': ["Total Assets"]}

alt_liabeq = {'Current Liabilities': ["Total Current Liabilities", "Current Liabilities"],
              'Long-term Debt': ["Long Term Debt", "Long-term Debt"],
              'Deposits': ["Deposits"],
              'Total Equity': ["Total Equity", "Shareholders' Equity", "Total Shareholders' Equity",
                               "Total stockholders' equity"],
              'Total Liabilities': ["Total Liabilities"],
              'Total Liabilities and Equity': ["Total Liabilities And Equity",
                                               "Total Liabilities And Shareholders' Equity",
                                               "Total Liabilities And Stockholders' Equity"]}

alt_incst = {'Cost of Goods Sold': ["Cost of Goods Sold", "Cost of Revenue", "Cost of Sales", "Cost of net revenues"],
             'Net Profit': ["Total Profit", "Total Income"],
             'Revenue': ["Total revenue", "Revenue"],
             'Net interest income':['Net interest income', 'Net-interest income'],
             'Non-interest income': ['Non-interest income', 'Noninterest income', 'Non interest income'],
             'Interest Expense': ["Interest Expense"],
             'EBIT': ["EBIT"],
             'EBITDA': ["EBITDA"]}

# Loop through and fill up indexes, indexes that cannot be found are given np.nan
for a in assets_index:
    a_labels = [s for s in label_list if any(xs.casefold() in s.casefold() for xs in alt_assets[a])]
    a_indexes = [label_list.index(n) for n in a_labels]

    if len(a_indexes) > 1:
        all_vals = {}
        for v in a_indexes:
            all_vals[v] = mergedfs.iat[v, 1]

        max_val = max(list(all_vals.values()))
        assets_index[a] = list(all_vals.keys())[list(all_vals.values()).index(max_val)]
    elif len(a_indexes) == 1:
        assets_index[a] = a_indexes[0]
    else:
        assets_index[a] = np.nan

# Insert check for total assets?

for le in liabeq_index:
    le_labels = [s for s in label_list if any(xs.casefold() in s.casefold() for xs in alt_liabeq[le])]
    le_indexes = [label_list.index(n) for n in le_labels]

    if len(le_indexes) > 1:
        all_vals = {}
        for v in le_indexes:
            all_vals[v] = mergedfs.iat[v, 1]

        max_val = max(list(all_vals.values()))
        liabeq_index[le] = list(all_vals.keys())[list(all_vals.values()).index(max_val)]
    elif len(le_indexes) == 1:
        liabeq_index[le] = le_indexes[0]
    else:
        liabeq_index[le] = np.nan

# Insert check for total liability and equity?

for i in incst_index:
    i_labels = [s for s in label_list if any(xs.casefold() in s.casefold() for xs in alt_incst[i])]
    i_indexes = [label_list.index(n) for n in i_labels]

    if len(i_indexes) > 1:
        all_vals = {}
        for v in i_indexes:
            all_vals[v] = mergedfs.iat[v, 1]

        max_val = max(list(all_vals.values()))
        incst_index[i] = list(all_vals.keys())[list(all_vals.values()).index(max_val)]
    elif len(i_indexes) == 1:
        incst_index[i] = i_indexes[0]
    else:
        incst_index[i] = np.nan

# mergedfs.to_excel(r'C:\Users\Navneeth\Documents\WORK\BankSME\Programming\banksme\outbound\mergedfs.xlsx', header=True,
#                   index=True, na_rep='NAN')

for i in assets.keys():
    i_pos = assets_index[i]
    if i == 'Total Assets':
        if not np.isnan(i_pos):
            assets[i] = mergedfs.iloc[i_pos, 1:].tolist()
        else:
            assets[i] = [np.nan] * (mergedfs.shape[1] - 1)
    else:
        if not np.isnan(i_pos):
            assets[i] = np.multiply(np.divide(mergedfs.iloc[i_pos, 1:].tolist(),
                                              mergedfs.iloc[assets_index['Total Assets'], 1:]),
                                    [100]*len(mergedfs.iloc[i_pos, 1:].tolist()))
        else:
            assets[i] = [np.nan] * (mergedfs.shape[1] - 1)

assets_df = pd.DataFrame.from_dict(assets, orient='index')
assets_df.insert(0, 'Labels', assets_df.index.values.tolist(), True)
assets_df.index = range(0, assets_df.shape[0])
assets_df.columns = mergedfs.columns

print()
print(assets_df)

for i in liabeq.keys():
    i_pos = liabeq_index[i]
    if i == 'Total Liabilities and Equity':
        if not np.isnan(i_pos):
            liabeq[i] = mergedfs.iloc[i_pos, 1:].tolist()
        else:
            liabeq[i] = [np.nan] * (mergedfs.shape[1] - 1)
    else:
        if not np.isnan(i_pos):
            liabeq[i] = np.multiply(np.divide(mergedfs.iloc[i_pos, 1:].tolist(),
                                              mergedfs.iloc[liabeq_index['Total Liabilities and Equity'], 1:]),
                                    [100]*len(mergedfs.iloc[i_pos, 1:].tolist()))
        else:
            liabeq[i] = [np.nan] * (mergedfs.shape[1] - 1)

liabeq_df = pd.DataFrame.from_dict(liabeq, orient='index')
liabeq_df.insert(0, 'Labels', liabeq_df.index.values.tolist(), True)
liabeq_df.index = range(0, liabeq_df.shape[0])
liabeq_df.columns = mergedfs.columns

print()
print(liabeq_df)

for i in incst.keys():
    i_pos = incst_index[i]
    if not np.isnan(i_pos):
        incst[i] = mergedfs.iloc[i_pos, 1:].tolist()
    else:
        incst[i] = [np.nan] * (mergedfs.shape[1] - 1)

incst_df = pd.DataFrame.from_dict(incst, orient='index')
incst_df.insert(0, 'Labels', incst_df.index.values.tolist(), True)
incst_df.index = range(0, incst_df.shape[0])
incst_df.columns = mergedfs.columns

print()
print(incst_df)
