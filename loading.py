from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, get_column_letter
import pandas as pd

def translate_school(school:  str) -> str:
    '''
    Translate school name to a standardized format
    
    Example:
    - 'St. John's' -> 'St.'
    - 'XV. gimnazija' -> 'XV.'
    - 'Gimnazija Kranj' -> 'GK'
    - 'Osnovna šola Ketteja in Murna' -> 'OSKIM'
    '''
    
    # If the school has a roman numeral, return the numeral
    if '.' in school:
        return school[0:school.index('.')+1]
    
    # Else, return the first letter of each word in the school name
    else:
        return ''.join([i[0].upper() for i in school.split(' ')])
    
def load_all_schools(sheet: object)  -> list[tuple[str, tuple[str]]]:
    '''
    Will load all schools from a passed sheet.
    
    Returns them in a list of tuples, where the first element is the school name and the second element is the range of the school name.
    
    Example:
    [('St.', ('A1', 'A2')), ('GK', ('A3', 'A5')), ('OSKIM', ('A6', 'A8'))]
    '''
    
    schools = []

    # Gets all the merged cell ranges, which contain the school names
    for merged_cell_range in sheet.merged_cells.ranges:
    
        # Get the range of the merged cell
        school_range = tuple(str(merged_cell_range).split(':'))
    
        # Translate the school name to a standardized format
        school_name = sheet[school_range[0]].value
        school_name = translate_school(school_name)
    
        # Append the school name and the range to the schools list in a tuple format
        schools.append((school_name, school_range))
    
    # Sort the schools by the row number of the first cell in the range (the left cell)
    schools = sorted(schools, key=lambda x: int(x[1][0][1:]))
        
    return schools

def load_classes(sheet: object, school_tuple: tuple[str, tuple[str]], next_school_start_row: int) -> list[tuple[str, str]]:
    '''
    Will load all classes from a passed sheet, for a passed school.
    Besides the sheet, the function requires a tuple containing the school name and the range of the school name. It also requires the row number of the next school so that it knows where to stop.
    Returns a list of tuples, where the first element is the class name and the second element is the start of the class range.
    
    Example:
    [('1. A', 'A2'), ('1. B', 'A3'), ('2. A', 'A4')]
    '''
    
    #  Get the school name and the start of the school range
    school_name, school_start = school_tuple[0], school_tuple[1][0]
    
    # Get the column and row of the cell where the school starts
    column, start_row = coordinate_from_string(school_start)
    start_row = int(start_row)
    
    classes = []
    
    # Iterate through the rows from the start of the school to the start of the next school
    for row in range(start_row+1, next_school_start_row):
        
        class_name = sheet[column+str(row)].value
        
        # If the cell is not empty, append the class name and the start of the class range to the classes list
        if class_name != None:
            classes.append((class_name, column+str(row)))
        
    return classes

def load_volunteers(sheet: object, location_name: str, school_name: str, class_tuple: tuple[str, str], next_class_start_row: int) -> pd.DataFrame:
    '''
    Will load all volunteers from a passed sheet, for a passed class (and school).
    
    Returns a DataFrame with the following columns:
    - volunteer_name
    - volunteer_dates
    - volunteer_class
    - location_name
    - volunteer_school
    '''
    
    
    volunteers_dict = {
        'volunteer_name': [],
        'volunteer_dates': [],
        'volunteer_class': [],
        'location_name': [],
        'volunteer_school': []
    }
    
    class_name = class_tuple[0]
    class_range = class_tuple[1]
    
    column, start_row = coordinate_from_string(class_range)
    start_row = int(start_row)
    
    column = get_column_letter(column_index_from_string(column)+2)
    
    for row in range(start_row, next_class_start_row):
        
        volunteer_name = sheet[column+str(row)].value
        
        matching_cells = [sheet[get_column_letter(cell.column)+DATES_ROW] for cell in sheet[row] if cell.value == 'da']
        matching_cells = [cell.value for cell in matching_cells if cell.value != None]
        
        if volunteer_name != None:
            volunteers_dict['volunteer_name'].append(volunteer_name)
            volunteers_dict['volunteer_dates'].append(matching_cells)
            volunteers_dict['volunteer_class'].append(class_name)
            volunteers_dict['volunteer_school'].append(school_name)
            volunteers_dict['location_name'].append(location_name)
        
    return pd.DataFrame(volunteers_dict)