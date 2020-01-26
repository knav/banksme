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
label_col = bsheet.iloc[:, 0]
label_list = label_col.tolist()

# Setup the dataframes for assets, liabeq and incst
assets_df = pd.DataFrame({'Labels': ['Current Assets', 'Accounts Receivable', 'Marketable Securities',
                                     'Loans', 'Cash', 'Inventory', 'PPE']})
liabeq_df = pd.DataFrame({'Labels': ['Current Liabilities', 'Long-term Debt', 'Deposits', 'Total Equity',
                                     'Total Liabilities']})
incst_df = pd.DataFrame({'Labels': ['Cost of Goods Sold', 'Net Profit', 'Revenue', 'Net interest income',
                                    'Non-interest income', 'Interest Expense', 'EBIT', 'EBITDA']})

# Dictionary to save relevant index of columns
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

incst = {'Cost of Goods Sold': 0,
         'Net Profit': 0,
         'Revenue': 0,
         'Net interest income': 0,
         'Non-interest income': 0,
         'Interest Expense': 0,
         'EBIT': 0,
         'EBITDA': 0}

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

alt_incst = {'Cost of Goods Sold': ["Cost of Goods Sold", "Cost of Revenue", "Cost of Sales", "Cost of net revenues"],
             'Net Profit': ["Net Profit", "Total Profit", "Net Income", "Total Income"],
             'Revenue': ["Total revenue", "Net revenue", "Revenue"],
             'Net interest income':['Net interest income', 'Net-interest income'],
             'Non-interest income': ['Non-interest income', 'Noninterest income', 'Non interest income'],
             'Interest Expense': ["Interest Expense"],
             'EBIT': ["EBIT"],
             'EBITDA': ["EBITDA"]}


