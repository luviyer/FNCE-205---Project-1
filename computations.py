import pandas as pd
import matplotlib.pyplot as plt

def create_df(excel_file, period):
    xl = pd.ExcelFile(excel_file)
    df = xl.parse('Values by Period')
    # Insert the initial investment into the dataframe in order to complete monthly returns.
    df.loc[-1] = ['12/31/{}'.format(period[0:4]), 10000, 10000, 10000]
    df.index = df.index + 1  # Shift the inserted row to the top
    df = df.sort_index()  # Sort by index

    # Compute the monthly returns from the values at the end of each period.
    df = df.fillna(method='ffill')
    df[['A1', 'A2', 'A3']] = df[['A1', 'A2', 'A3']]/df[['A1', 'A2', 'A3']].shift(1) - 1
    # Remove the Market data and the first row.
    df = df.loc[1:, ['Period', 'A1', 'A2', 'A3']]

    df['Period'] = pd.to_datetime(df['Period'], format='%m/%d/%Y')
    df.set_index('Period', inplace=True)
    df = df.rename(columns={'A1': 'High Cap', 'A2': 'Mid Cap', 'A3': 'Low Cap'})
    return df

# Create a bunch of dataframes for the n values in a time period
# and put them into a multi-index frame
n_values = [1, 2, 4, 10, 20, 30, 40, 70, 100]
ten_year_period = ['20062016']
pn = pd.DataFrame()

total_returns = {}
total_stats = {}
total_corr = {}
# Make a separate frame for each period
for period in ten_year_period:
    for n in n_values:
        excel_file = '{}/{}Stock{}.xls'.format(period, n, period)
        df = create_df(excel_file, period)
        add_1 = df + 1
        cum_prod = add_1.cumprod()
        data = [{'ar_mean': df['High Cap'].mean(), 'std': df['High Cap'].std(),
                 'geo_mean': cum_prod.iloc[-1]['High Cap']**(1.0/df.shape[0]) - 1},
                {'ar_mean': df['Mid Cap'].mean(), 'std': df['Mid Cap'].std(),
                 'geo_mean': cum_prod.iloc[-1]['Mid Cap']**(1.0/df.shape[0]) - 1},
                {'ar_mean': df['Low Cap'].mean(), 'std': df['Low Cap'].std(),
                 'geo_mean': cum_prod.iloc[-1]['Low Cap']**(1.0/df.shape[0]) - 1}]

        stats = pd.DataFrame(data, index=['High Cap', 'Mid Cap', 'Low Cap'])
        corr_matrix = df[['High Cap', 'Mid Cap', 'Low Cap']].corr()

        total_returns['N={}'.format(n)] = df
        total_stats['N={}'.format(n)] = stats
        total_corr['N={}'.format(n)] = corr_matrix

    total_returns = pd.concat(total_returns.values(), keys=total_returns.keys())
    total_stats = pd.concat(total_stats.values(), keys=total_stats.keys())
    total_corr = pd.concat(total_corr.values(), keys=total_corr.keys())

print(total_returns)
print(total_stats)
print(total_corr)

