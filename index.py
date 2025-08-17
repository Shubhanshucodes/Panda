import pandas as pd
import numpy as np
import os
from datetime import datetime
import xlrd

input_folder = 'shriram_input/'
output_folder = 'shriram_output/'


def read_value_from_excel(filename, sheet_name, column="C", row=4):
    # print(filename)
    """Read a single cell value from an Excel file"""
    return pd.read_excel(filename, sheet_name=sheet_name, skiprows=row - 1, usecols=column, nrows=1, header=None,
                         names=["Value"]).iloc[0]["Value"]


rows = ["FUND_ID", "PORTFOLIO_AS_ON", "ISIN", "STOCK_NAME", "INDUSTRY_RATING", "QUANTITY", "%_to_NAV", "MARKET_VALUE",
        "LAST_UPD_TMS"]
rows_names = {'ISIN': 'ISIN', 'Name of the Instrument / Issuer': 'STOCK_NAME', 'Rating / Industry^': 'INDUSTRY_RATING',
              'Quantity': 'QUANTITY', 'Market value\n(Rs. in Lakhs)': 'MARKET_VALUE', '% to AUM': '%_to_NAV'}


scheme_master = pd.read_csv("SCHEME_MASTER.csv")
scheme_master = scheme_master.dropna(subset=['SCHEME_CODE', 'Sheet Name'])
sheetname_to_schemecode = scheme_master.set_index('Sheet Name')['SCHEME_CODE'].to_dict()
sheetname_to_schemename = scheme_master.set_index('Sheet Name')['SCHEME_NAME'].to_dict()
#print(sheetname_to_schemename, 'hereitis')

with pd.ExcelFile('shriram_input/Monthly-Portfolio-Shriram-Mutual-Fund-January-2024.xls') as excel_file:
    sheet_names = excel_file.sheet_names
    sheet_names = [sheet_name for sheet_name in sheet_names if sheet_name != 'INDEX']

    for sheet_name in sheet_names:
        print(sheet_name)
        # Read the sheet into a DataFrame
        # cheme = excel_file.parse(sheet_name)

        # for sheet_name in pd.ExcelFile(filename).sheet_names:
        #    if sheet_name in ['Common Notes', 'Scheme']:
        #        continue
        try:
            portfolio_as_on = read_value_from_excel(excel_file, sheet_name, 'B', 3)[16:]
            print(portfolio_as_on)

            df = pd.read_excel(excel_file, sheet_name, header=20)
            # print(sheet_name)
            print(df.columns)
            if 'Unnamed: 0' in df.columns:
                df.drop(columns=['Unnamed: 0'], inplace=True)
            if 'Unnamed: 7' in df.columns:
                df.drop(columns=['Unnamed: 7'], inplace=True)
            if 'Coupon (%)' in df.columns:
                df.drop(columns=['Coupon (%)'], inplace=True)
            if 'YTM~' in df.columns:
                df.drop(columns=['YTM~'], inplace=True)
            if 'YTC^' in df.columns:
                df.drop(columns=['YTC^'], inplace=True)
            if 'null' in df.columns:
                df.drop(columns=['null'], inplace=True)
            if 'ESG Score' in df.columns:
                df.drop(columns=['ESG Score'], inplace=True)
            if 'Yield' in df.columns:
                df.drop(columns=['Yield'],
                        inplace=True)
            print(df.columns)
            # if '% to AUM' in df.columns:
            # df.drop(columns=['% to AUM'], inplace=True)
            df.columns = ['ISIN','Name of the Instrument / Issuer','Rating / Industry^', 'Quantity',
                          'Market value\n(Rs. in Lakhs)', '% to AUM']

            # Selecting required columns
            df_final = df[rows_names.keys()]
            df_final.columns = [rows_names[column] for column in df_final.columns if column in rows_names]
            df_final = df_final.dropna(axis='rows', how='all')
            df_final = df_final[(df_final['ISIN'].str.startswith('IN') | df_final['ISIN'].str.startswith('DB'))]
            # Adding rows with same values
            df_final['PORTFOLIO_AS_ON'] = portfolio_as_on
            df_final['FUND_ID'] = sheetname_to_schemecode.get(sheet_name.strip(), None)
            # if df_final['FUND_ID'].iloc[0] is not None:
            # print(df_final['FUND_ID'])
            df_final['QUANTITY'] = pd.to_numeric(df_final['QUANTITY'], errors='coerce').fillna(0)
            df_final['%_to_NAV'] = pd.to_numeric(df_final['%_to_NAV'], errors='coerce').fillna(0)
            df_final['%_to_NAV'].fillna(0,inplace=True)
            df_final['MARKET_VALUE'] = pd.to_numeric(df_final['MARKET_VALUE'], errors='coerce').fillna(0)

            df_final['LAST_UPD_TMS'] = None
            # df_final[rows].to_excel(scheme[sheet_name.strip()]+".xlsx",index=False)
            # print('LAST_UPD_TMS')
            # print(scheme[sheet_name.strip()])
            # print('LAST_UPD_TMS1')
            df_final['%_to_NAV'] = pd.to_numeric(df_final['%_to_NAV'], errors='coerce')
            df_final['%_to_NAV'] = (df_final['%_to_NAV'] * 100).round(2)
            #print(df_final)
            out_path = 'C:/Users/Ankur/PycharmProjects/Exel Input New/shriram_output/' + sheetname_to_schemename[
                sheet_name.strip()] + ".csv"
            print(out_path)
            # df_final[rows].to_excel(out_path + scheme[sheet_name.strip()] + ".csv", index=False)
            df_final[rows].to_csv(out_path, index=False)

        except Exception as e:
            print("Exception occurred for value '" + sheet_name + "': " + repr(e))