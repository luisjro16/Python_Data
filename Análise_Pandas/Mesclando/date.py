import pandas as pd

date_range = pd.date_range(start='20230101', end='20230328')

for date in date_range:
    print(date)