### imports and load data


```python
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
```




<div><div id=d99ae80b-f534-4df9-af2e-47f610f0e6b4 style="display:none; background-color:#9D6CFF; color:white; width:200px; height:30px; padding-left:5px; border-radius:4px; flex-direction:row; justify-content:space-around; align-items:center;" onmouseover="this.style.backgroundColor='#BA9BF8'" onmouseout="this.style.backgroundColor='#9D6CFF'" onclick="window.commands?.execute('create-mitosheet-from-dataframe-output');">See Full Dataframe in Mito</div> <script> if (window.commands?.hasCommand('create-mitosheet-from-dataframe-output')) document.getElementById('d99ae80b-f534-4df9-af2e-47f610f0e6b4').style.display = 'flex' </script> <table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>id</th>
      <th>DoB</th>
      <th>DoT</th>
      <th>amt</th>
      <th>province</th>
      <th>Pensionable Service***:</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>1</td>
      <td>1996-02-24 00:00:00</td>
      <td>1997-01-30</td>
      <td>123.45</td>
      <td>Ontario</td>
      <td>1.04650</td>
    </tr>
    <tr>
      <th>1</th>
      <td>2</td>
      <td>1997-05-22 00:00:00</td>
      <td>2003-08-19</td>
      <td>1234.34</td>
      <td>Quebec</td>
      <td>22.44558</td>
    </tr>
    <tr>
      <th>2</th>
      <td>3</td>
      <td>1999-04-17 00:00:00</td>
      <td>1997-08-09</td>
      <td>23234.56</td>
      <td>Nova Scotia</td>
      <td>9.75409</td>
    </tr>
  </tbody>
</table></div>



### assignments and functions


```python
lookup = {1: 'janvier', 2: 'février', 3: 'mars', 4: 'avril', 5: 'mai',
          6: 'juin', 7: 'juillet', 8: 'août', 9: 'septembre', 10: 'octobre', 11: 'novembre', 12: 'décembre'}

fr_prov_dict = {'British Columbia': 'Colombie-Britannique',
                'Did Not Work in ON or NS': "N'a pas travaillé en Ontario ou en Nouvelle-Écosse",
                "New Brunswick": 'Nouveau-Brunswick',
                'Nova Scotia': 'Nouvelle-Écosse',
                'Quebec': 'Québec',
                'Newfoundland and Labrador': 'Terre-Neuve-et-Labrador',
                'Newfoundland':'Terre-Neuve-et-Labrador',
                'Prince Edward Island': 'Île-du-Prince-Édouard',
                'Northwest Territories': 'Les Territoires du Nord-Ouest',
                'Defined Contribution': 'Cotisations Déterminée'
                }


def translate_provinces(t_prov_str):
    for key in fr_prov_dict.keys():
        t_prov_str = t_prov_str.replace(key, fr_prov_dict[key])

    return t_prov_str


def conv_date(t_date):
    if type(t_date) != datetime.datetime and type(t_date) != pd._libs.tslibs.timestamps.Timestamp:
        return ""
    else:
        return_str = str(t_date.day) + ' ' + \
            lookup[t_date.month] + ' ' + str(t_date.year)
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
            return (round((x)/multiplier, position))
        if num < 0:
            x = math.floor(num)
            return (round((x)/multiplier, position))
    else:
        return (round(num/multiplier, position))


def conv_curr(t_amt, dec_places=2, dollar_sign=True):
    t_str = str(fix_round(t_amt, dec_places))
    t_list = t_str.split('.')

    before_dec = t_list[0]

    before_dec_list = [before_dec[0:len(before_dec) % 3]] + [before_dec[i:i+3]
                                                             for i in range(len(before_dec) % 3, len(before_dec), 3)]

    before_dec_str = ' '.join(before_dec_list)

    while len(t_list[1]) < dec_places:
        t_list[1] += '0'

    if dollar_sign == True:
        return_str = before_dec_str + ',' + t_list[1] + ' $'
    else:
        return_str = before_dec_str + ',' + t_list[1]
    return return_str

```

### user input


```python
curr_col_list_input = ['amt']
date_col_list_input = ['DoB', 'DoT']
prov_col_list_input = ['province']
round_to_dec_list_input = ['Pensionable Service***:']
```

### run and output


```python
for col in round_to_dec_list_input:
    df[col] = df[col].fillna(0)
    df[col] = df[col].apply(conv_curr, dec_places=4, dollar_sign=False)
    df[col] = df[col].apply(lambda x: x.strip())

for col in curr_col_list_input:
    df[col] = df[col].fillna(0)
    df[col] = df[col].apply(conv_curr)
    df[col] = df[col].apply(lambda x: x.strip())

for col in date_col_list_input:
    df[col] = df[col].apply(conv_date)

for col in prov_col_list_input:
    df[col] = df[col].fillna("")
    df[col] = df[col].apply(translate_provinces)
    

df.to_excel('test_out.xlsx', index=False)

# files.download('test_out.xlsx')

df
```

    C:\Users\ronal\AppData\Local\Temp/ipykernel_2004/4088978067.py:44: DeprecationWarning: `np.float` is a deprecated alias for the builtin `float`. To silence this warning, use `float` by itself. Doing this will not modify any behavior and is safe. If you specifically wanted the numpy scalar type, use `np.float64` here.
    Deprecated in NumPy 1.20; for more details and guidance: https://numpy.org/devdocs/release/1.20.0-notes.html#deprecations
      if type(num) != float and type(num) != int and type(num) != numpy.float:
    




<div><div id=46fc7157-7560-4e65-bf10-415c4d91a614 style="display:none; background-color:#9D6CFF; color:white; width:200px; height:30px; padding-left:5px; border-radius:4px; flex-direction:row; justify-content:space-around; align-items:center;" onmouseover="this.style.backgroundColor='#BA9BF8'" onmouseout="this.style.backgroundColor='#9D6CFF'" onclick="window.commands?.execute('create-mitosheet-from-dataframe-output');">See Full Dataframe in Mito</div> <script> if (window.commands?.hasCommand('create-mitosheet-from-dataframe-output')) document.getElementById('46fc7157-7560-4e65-bf10-415c4d91a614').style.display = 'flex' </script> <table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>id</th>
      <th>DoB</th>
      <th>DoT</th>
      <th>amt</th>
      <th>province</th>
      <th>Pensionable Service***:</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>1</td>
      <td>24 février 1996</td>
      <td>30 janvier 1997</td>
      <td>123,45 $</td>
      <td>Ontario</td>
      <td>1,0465</td>
    </tr>
    <tr>
      <th>1</th>
      <td>2</td>
      <td>22 mai 1997</td>
      <td>19 août 2003</td>
      <td>1 234,34 $</td>
      <td>Québec</td>
      <td>22,4456</td>
    </tr>
    <tr>
      <th>2</th>
      <td>3</td>
      <td>17 avril 1999</td>
      <td>9 août 1997</td>
      <td>23 234,56 $</td>
      <td>Nouvelle-Écosse</td>
      <td>9,7541</td>
    </tr>
    <tr>
      <th>3</th>
      <td>4</td>
      <td></td>
      <td>9 août 1999</td>
      <td>3 434,34 $</td>
      <td>Colombie-Britannique</td>
      <td>3,0525</td>
    </tr>
    <tr>
      <th>4</th>
      <td>5</td>
      <td></td>
      <td>15 février 2002</td>
      <td>45,30 $</td>
      <td>Ontario (Cotisations Déterminée)</td>
      <td>19,9480</td>
    </tr>
    <tr>
      <th>5</th>
      <td>6</td>
      <td></td>
      <td></td>
      <td>1,00 $</td>
      <td>Québec (Cotisations Déterminée)</td>
      <td>4,0810</td>
    </tr>
    <tr>
      <th>6</th>
      <td>7</td>
      <td></td>
      <td></td>
      <td>0,00 $</td>
      <td>Nouvelle-Écosse (Cotisations Déterminée)</td>
      <td>20,3073</td>
    </tr>
    <tr>
      <th>7</th>
      <td>8</td>
      <td></td>
      <td></td>
      <td>0,00 $</td>
      <td>Colombie-Britannique (Cotisations Déterminée)</td>
      <td>555 555,7362</td>
    </tr>
    <tr>
      <th>8</th>
      <td>9</td>
      <td></td>
      <td></td>
      <td>0,00 $</td>
      <td></td>
      <td>0,0000</td>
    </tr>
  </tbody>
</table></div>


