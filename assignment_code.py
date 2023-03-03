# -*- coding: utf-8 -*-
"""
Created on Wed Jan 11 21:48:27 2023

@author: Ishan Jain
"""

# importing libraries
import pandas as pd
import time

# class definition
class statGather:
    def __init__(self):
        self.input_file = "Python Test 1 - billings_europe.csv"

    # method to read the data and perform cleanup and transformations
    def cleanup(self):
        df = pd.read_csv(self.input_file, skiprows=3, index_col=[0], low_memory=False)
        # dropping the null column from data
        df = df.drop('2', axis = 1)
        # creating a df of the data header and transposing the dataframe
        df1 = df.head(3)
        df1 = df1.T
        # Renaming the column names of header dataframe
        df1.columns = ['Segment', 'Type', 'Subtype']
        # using forward fill to fill up the missing Type with previous non-missing value
        df1['Type'] = df1['Type'].ffill()
        # splitting the Segment into Segment and Period
        df1[['Segment','Period']] = df1['Segment'].str.split(' - ',expand=True)
        # empty list to store processed rows
        df_lst = []

        # iterarting over data rows
        for idx, row in df.iloc[3:].iterrows():
            # checking if the row is null
            check_len = list(set(list(row.isna())))
            if len(check_len) > 1:
                x = df1.copy()
                x['Date'] = idx
                x['Value'] = row
                df_lst.append(x)

        # createing a dataframe from all the processed rows           
        fin_df = pd.concat(df_lst)
        # converting the data type of date column to "Date" type
        fin_df['Date'] = pd.to_datetime(fin_df['Date'], format='%d-%b-%y')
        # converting the data type of value column to "Float" type
        fin_df['Value'] = fin_df['Value'].astype('float64')
        fin_df['Value'] = fin_df['Value'].fillna(0)
        
        
        # calculating sum of billing by country in country_sum dataframe
        country_data = fin_df[fin_df['Type']=="Countries"]
        country_sum = country_data.groupby(['Subtype']).sum('Value').reset_index()
        country_sum.columns = ['Countries','Billings']

        # calculating sum of billing by Period in period_data dataframe        
        period_data = fin_df[(fin_df['Type']=="Market") & (fin_df['Date']>='2016-01-01')]
        period_data = period_data.groupby(['Period']).sum('Value').reset_index()
        period_data.columns = ['Period','Billings']

        # Calculatuing Segment summary stats in segment_data dataframe
        segment_data = fin_df.groupby('Segment')['Value'].agg(['sum','mean','median','min','max','skew','count']).reset_index()
        
        # writing the output in excel
        with pd.ExcelWriter('Python Test 1 - Python Exercise Output.xlsx', engine='xlsxwriter') as writer:
            country_sum.to_excel(writer, sheet_name='Output', index = False, startrow=3, startcol=0)
            period_data.to_excel(writer, sheet_name='Output', index = False, startrow=3, startcol=4)
            segment_data.to_excel(writer, sheet_name='Segment Summary Stats', index = False, startrow=0, startcol=0)


if __name__ == "__main__":
    start_time = time.time()
    obj = statGather()
    obj.cleanup()
    end_time = time.time()
    print(f"Total Time elapsed in execution: {end_time - start_time} seconds.")


