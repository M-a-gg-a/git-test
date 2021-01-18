# -*- coding: utf-8 -*-
"""
Created on Mon Jan 11 10:07:38 2021

@author: Kolovos
"""

# ----------'Reading' from ecxel 'EconomicAnalysisM3.xlsx'-------------------
# pandas can READ formulas Directly from excel sheet
import pandas as pd

sheet = pd.read_excel('EconomicAnalysisM3.xlsx') 

# assign values to var in [$/MWh] from 'EconomicAnalysisM3.xlsx'
C26 = sheet.iloc[24, 2] # .iloc[] picks specific cell form xlsx
D27 = sheet.iloc[25, 3] # .iloc[] picks specific cell form xlsx
E26 = sheet.iloc[24, 4] # .iloc[] picks specific cell form xlsx
E27= sheet.iloc[25, 4]  # .iloc[] picks specific cell form xlsx