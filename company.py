import pandas as pd
import numpy as np


class Company(object):
    def __init__(self, year, assets, liabeq, incst):
        self.year = year
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

    def updatevals(self):
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
            'Total debt/Total assets': self.liabeq['Total Liabilities'] / 100,
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

    def getindustry(self):
        # Accounts Receivable
        if self.assets['Accounts Receivable'] == np.nan:
            self.scores['Bank'] += 5
            self.tags.append("No accounts receivable field")

        # Inventory
        if self.assets['Inventory'] == np.nan:
            self.scores['Bank'] += 5
            self.tags.append("No inventory field")
        elif self.assets['Inventory'] < 2:
            self.scores['Bank'] += 3
            self.scores['SNS'] += 5

        # Loans - Bank exclusive
        if self.assets['Loans'] >= 50:
            self.scores['Bank'] += 5
            self.tags.append("Loans form majority of assets")

        # Deposits - Bank exclusive
        if self.liabeq['Deposits'] >= 50:
            self.scores['Bank'] += 5
            self.tags.append("Deposits form majority of liabilities")

        # Cash and Marketable Securities
        if self.assets['Cash'] + self.assets['Marketable Securities'] <= 20:
            self.scores['F&B'] += 3
        elif self.assets['Cash'] + self.assets['Marketable Securities'] >= 40:
            self.scores['SNS'] += 3

        # Inventory
        if self.assets['Inventory'] <= 10:
            self.scores['SNS'] += 3
            self.scores['Online Mkt'] += 3
            self.tags.append("Low inventory levels")

        # Inventory turnover
        if 1 < (365 / self.ratios['Inventory turnover']) <= 10:
            self.scores['F&B'] += 5
            if (365 / self.ratios['Inventory turnover']) <= 5:
                self.tags.append("Very rapid inventory turnover, likely to be a fast-food joint")
            else:
                self.tags.append("Rapid inventory turnover, dealing with perishables")
        elif 10 < (365 / self.ratios['Inventory turnover']) <= 14:
            self.scores['F&B'] += 3
            self.scores['Grocery'] += 3
            self.tags.append("Fairly fast inventory turnover, possibly dealing with perishables")

        # PPE
        if self.assets['PPE'] >= 70:
            self.scores['Construction'] += 3
            self.tags.append("High PPE, indicating capital-intensive business")
        elif 40 <= self.assets['PPE'] < 70:
            self.scores['F&B'] += 4
            self.scores['Grocery'] += 3
            self.tags.append("Fairly high PPE, rather capital-intensive")
        elif 15 <= self.assets['PPE'] < 30:
            self.scores['SNS'] += 3
            self.tags.append("Low PPE")
        elif self.assets['PPE'] <= 5:
            self.scores['Online Mkt'] += 3

        # Total Equity
        if self.liabeq['Total Equity'] >= 45:
            self.scores['SNS'] += 3

        # EBITDA Margin

        #
        print(pd.DataFrame.from_dict(self.ratios, orient="index", columns=['Ratios']))
        print()
        print(pd.DataFrame.from_dict(self.scores, orient="index", columns=['Scores']))
        print()
        print(self.tags)
