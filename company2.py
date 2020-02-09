import pandas as pd
import numpy as np

class Company2(object):
    def __init__(self, assets, liabeq, incst):
        self.assets = assets
        self.liabeq = liabeq
        self.incst = incst
        self.ratios = {
            'Current assets/Current Liabilities': self.assets['Current Assets'] / self.liabeq['Current Liabilities'],
            'Cash, marketable securities and accounts receivable/current liabilities':
                (self.assets['Cash']
                    + self.assets['Accounts Receivable']
                    + self.assets['Marketable Securities'])
                / self.liabeq['Current Liabilities'],
            'Inventory turnover':
                (self.incst['Cost of Goods Sold'] / (self.assets['Inventory']*self.assets['Total Assets']))*100,
            'Receivables collection period (days)': 365 / (self.incst['Revenue'] / self.assets['Accounts Receivable']),
            'Total debt/Total assets': self.liabeq['Total Liabilities'] / self.assets['Total Assets'],
            'Long-term debt/Capitalization': None,
            'Revenue/Total Assets': self.incst['Revenue'] / self.assets['Total Assets'],
            'Net profit/Revenue': self.incst['Net Profit'] / self.incst['Revenue'],
            'Net Profit/Total Assets': self.incst['Net Profit'] / self.assets['Total Assets'],
            "Total assets/Shareholders' Equity": 100 / self.liabeq['Total Equity'],
            "Net profit/Shareholders' Equity":
                self.incst['Net Profit']
                / ((self.liabeq['Total Equity']/100)*self.liabeq['Total Liabilities and Equity']),
            'EBIT/Interest expense': self.incst['EBIT'] / self.incst['Interest Expense'],
            'EBITDA/Revenue': self.incst['EBITDA'] / self.incst['Revenue']}
        self.industry = ""
        self.scores = {'Bank': 0,
                       'F&B': 0,
                       'SNS': 0,
                       'Online Mkt': 0,
                       'Retail': 0,
                       'Grocery': 0,
                       'Contruction': 0}
        self.tags = []