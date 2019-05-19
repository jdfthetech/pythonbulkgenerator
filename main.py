import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import numpy as np

print('Welcome to the Bulk File Generator,')
print('The classy way to build your Amazon bulk file.\n')
print('!!!---------------------IMPORTANT---------------------!!!')  
print('It is imperative that your input file contains one column of SKUs (nonduplicated) with a header,')    
print('that it is named input.xlsx and has a sheetname of Sheet1.\n')    

campaignName = input("Please input your desired Campaign Name and hit enter: \n")
bidPlus = input("Enter 'Enabled' to turn on Bid+ otherwise just hit enter: \n")
maxBid = input("Please enter a numeric maximum per click bid (ex. 0.35'):\n")    
dailyBudget = input("Please input your daily budget (ex. 5.00'):\n") 


df = pd.read_excel('input.xlsx', sheet_name='Sheet1')

# create duplicates of skus
df2 =pd.DataFrame(np.repeat(df.values,2,axis=0))

#make a new dataframe to tear apart
df3 = df2




df3.insert(0, 'Campaign Name', campaignName, allow_duplicates = False)
df3.insert(1, 'Campaign Daily Budget', '', allow_duplicates = False)
df3.insert(2, 'Campaign Start Date', '', allow_duplicates = False)
df3.insert(3, 'Campaign End Date', '', allow_duplicates = False)
df3.insert(4, 'Campaign Targeting Type', '', allow_duplicates = False)
df3.insert(5, 'Ad Group Name', '', allow_duplicates = False)
df3.insert(6, 'Max Bid', '', allow_duplicates = False)
df3.insert(7, 'SKU', '', allow_duplicates = False)
df3.insert(8, 'Keyword', '', allow_duplicates = False)    
df3.insert(9, 'Match Type', '', allow_duplicates = False)      
df3.insert(10, 'Campaign Status', '', allow_duplicates = False)     
df3.insert(11, 'Ad Group Status', '', allow_duplicates = False) 
df3.insert(12, 'Status', '', allow_duplicates = False) 
df3.insert(13, 'Bid+', bidPlus, allow_duplicates = False) 

df3.ix[0,'Campaign Daily Budget'] = dailyBudget
df3.ix[0,'Campaign Status'] = 'Enabled'
df3.ix[0,'Campaign Targeting Type'] = 'Auto'





i = 1
while i < len(df2.index): 
    if (i % 2 == 0):
        df3.ix[i,'SKU'] = ""
        df3.ix[i,'Ad Group Name'] = df2.ix[i-1,0]
        df3.ix[i,'Ad Group Status'] = 'Enabled'
        df3.ix[i,'Max Bid'] = maxBid
        i = i+1
    else:
        df3.ix[i,'SKU'] = df2.ix[i-1,0]
        df3.ix[i,'Ad Group Name'] = df2.ix[i-1,0]
        df3.ix[i,'Status'] = 'Enabled'
        i = i+1       


# remove junk column
df3.drop(df3.columns[14], inplace=True, axis=1)

df3.to_excel('newCampaign.xlsx', index=False)

print('Your Campaign has been built!')
print('You may find the file in this directory entitled newCampaign.xlsx\n')
print('Please remember to add your campaign start and end dates\n')

