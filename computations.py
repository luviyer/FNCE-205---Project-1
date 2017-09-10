import pandas as pd
import matplotlib.pyplot as plt

def create_df(excel_file):
    xl = pd.ExcelFile(excel_file)
    df = xl.parse('Returns by Period')
    df['Period'] = pd.to_datetime(df['Period'], format='%m/%d/%Y')
    df.set_index('Period', inplace=True)

    return df

# Create a bunch of dataframes for the n values in a time period
# and put them into a panel
n_values = [1, 2, 4, 10, 20, 30, 40, 70, 100]
ten_year_period = [20062016]
pn = pd.DataFrame()

total_returns = {}
total_stats = {}
total_corr = {}
for period in ten_year_period:
    for n in n_values:
        excel_file = '{}/{}Stock{}.xls'.format(period, n, period)
        df = create_df(excel_file)
        add_1 = df + 1
        cum_prod = add_1.cumprod()
        data = [{'ar_mean': df['A1'].mean(), 'std': df['A1'].std(),
                 'geo_mean': cum_prod.iloc[-1]['A1']**(1.0/df.shape[0]) - 1},
                {'ar_mean': df['A2'].mean(), 'std': df['A2'].std(),
                 'geo_mean': cum_prod.iloc[-1]['A2']**(1.0/df.shape[0]) - 1},
                {'ar_mean': df['A3'].mean(), 'std': df['A3'].std(),
                 'geo_mean': cum_prod.iloc[-1]['A3']**(1.0/df.shape[0]) - 1}]

        stats = pd.DataFrame(data, index=['A1', 'A2', 'A3'])
        corr_matrix = df[['A1', 'A2', 'A3']].corr()

        total_returns['N={}'.format(n)] = df
        total_stats['N={}'.format(n)] = stats
        total_corr['N={}'.format(n)] = corr_matrix

    total_returns = pd.concat(total_returns.values(), keys=total_returns.keys())
    total_stats = pd.concat(total_stats.values(), keys=total_stats.keys())
    total_corr = pd.concat(total_corr.values(), keys=total_corr.keys())

print(total_returns)
print(total_stats)
print(total_corr)

