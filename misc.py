import pandas as pd

print(pd.to_datetime("18991231"))

last_day_of_previous_month = (pd.to_datetime("today") - pd.tseries.offsets.MonthEnd(1))#.strftime("%m/%d/%Y")
print(last_day_of_previous_month.strftime("%m.%y"))