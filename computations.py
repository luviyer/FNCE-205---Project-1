import pandas as pd

def create_df(excel_file):
    xl = pd.ExcelFile(excel_file)
    df = xl.parse('Returns by Period')
    df['Period'] = pd.to_datetime(df['Period'], format='%m/%d/%Y')
    df.set_index('Period', inplace=True)

    return df
excel_file = '1Stock20062016.xls'

df = create_df(excel_file)

print(df.head())
