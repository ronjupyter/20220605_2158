```python

# %% [markdown]
# ### imports and load data

# %%
import pandas as pd
import numpy
import datetime
import math

# import io
# from google.colab import files
# uploaded = files.upload()
# df = pd.read_excel(io.BytesIO(uploaded['<filename>.xlsx']))

df = pd.read_excel('test_data.xlsx', sheet_name="aio")

df.head(3)

# %% [markdown]
# ### assignments and functions

# %%
lookup = {1: 'janvier', 2: 'février', 3: 'mars', 4: 'avril', 5: 'mai',
            6: 'juin', 7: 'juillet', 8: 'août', 9: 'septembre', 10: 'octobre', 11: 'novembre', 12: 'décembre'}


def conv_date(t_date):
    if type(t_date) != datetime.datetime and type(t_date) != pd._libs.tslibs.timestamps.Timestamp:
        return ""
    else:
        return_str = str(t_date.day) + ' ' + lookup[t_date.month] + ' ' + str(t_date.year)
        return return_str


def fix_round(num, position=0):
    '''fixes rounding error in py - may not be accurate w/ extremely precise numbers such as 0.49999999999999998

	params:
	num (float/int)
	position (int): default is 0

	returns rounded result
    
    '''

    if type(num) != float and type(num) != int and type(num) != numpy.float:
        return 0.00

    multiplier = 10**position
    num = multiplier * num

    try:
        after_decimal = str(num).split('.')[1]
    except:
        return (num/multiplier)

    if after_decimal == '5':
        if num > 0:
            x = math.ceil(num) 
            return (round((x)/multiplier,position))
        if num < 0:
            x = math.floor(num) 
            return (round((x)/multiplier,position))
    else:
        return (round(num/multiplier, position))


def conv_curr(t_amt):
    t_str = str(fix_round(t_amt, 2))
    t_list = t_str.split('.')

    before_dec = t_list[0]
    
    before_dec_list = [before_dec[0:len(before_dec)%3]] + [before_dec[i:i+3] for i in range(len(before_dec)%3, len(before_dec), 3)]

    before_dec_str = ' '.join(before_dec_list)
    
    if len(t_list[1]) == 1:
        t_list[1] += '0' 

    return_str = before_dec_str + ',' + t_list[1] + ' $' 
    return return_str  

# %% [markdown]
# ### user input

# %%
curr_col_list_input = ['amt']
date_col_list_input = ['DoB', 'DoT']


for col in curr_col_list_input:
    df[col] = df[col].fillna(0)
    df[col] = df[col].apply(conv_curr)
    df[col] = df[col].apply(lambda x: x.strip())

for col in date_col_list_input:
    df[col] = df[col].apply(conv_date)


# %% [markdown]
# ### output

# %%
df.to_excel('test_out.xlsx', index=False)

# files.download('test_out.xlsx')

df

```