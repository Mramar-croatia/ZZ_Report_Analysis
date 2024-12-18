from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, get_column_letter
import pandas as pd
import datetime

from filtering import *
from loading import *

PATH = 'SAMPLES\zz_lista_volontera_23_24.xlsx'

if __name__ == '__main__': 

    year_df = load_year(PATH)
    year_df = remove_non_wanted_dates(year_df)
    
    year_df = year_df.sort_values(by='volunteer_dates', key=lambda x: x.apply(len), ascending=False)
    year_df = year_df.reset_index(drop=True)
    
    year_df['volounteer_count'] = year_df['volunteer_dates'].apply(len)
    
    # print(year_df['volounteer_count'].sum())
    print(year_df.head(15))