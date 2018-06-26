# -*- coding: utf-8 -*-
"""
Created on Sun Jun 24 07:19:14 2018

@author: Eric
"""

import os
import pandas as pd


my_path = r'''C:\Users\Eric\pytest'''  #path with my files

files = os.listdir(path=my_path)  #creates list of files with my directory

sheets_expanded = []

#creates list of Excel Files--anything that ends in "xlsx"
for filename in files:
    if filename.endswith("xlsx"):
        sheets_expanded.append(filename)
        
print(sheets_expanded)  #prints this list; test to see what files the list has


os.chdir(path=my_path)  #changes acvite directory to diretory with my files


all_data = pd.DataFrame()

#creates Dataframe with all data from the Excel files
for f in sheets_expanded:
    df=pd.read_excel(f)
    all_data = all_data.append(df)
    
print(all_data.head())  #Prints first 5 rows of data; test

new_path = r'''C:\Users\Eric\Documents\Output'''

os.chdir(path = new_path) #change active directory

all_data.to_excel('Test_Combine_Excel.xlsx',sheet_name='NewSheet') #saves data as Excel file


