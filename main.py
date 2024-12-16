from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, get_column_letter
import pandas as pd

from loading import *

if __name__ == '__main__':
    
    # Load the workbook, data_only=True to get the values of the formulas not the formulas themselves
    wb = load_workbook('SAMPLES\zz_lista_volontera_24_25.xlsx', data_only=True)

    sheets = wb.sheetnames
    
    location_sheet = wb[sheets[1]]
    
    location_df = load_location(wb, sheets[1])
    
    print(location_df)