from datetime import datetime

def remove_dates_for_person(tuples_list, cutoff):
    return [
        (location, date) for location, date in tuples_list
        if datetime.strptime(date, "%d.%m.%Y.") <= cutoff
    ]

def remove_non_wanted_dates(year_df, cutoff_date):
    
    # "25.12.2023."
    
    cutoff_date_obj = datetime.strptime(cutoff_date, "%d.%m.%Y.")
    year_df['volunteer_dates'] = year_df['volunteer_dates'].apply(remove_dates_for_person, cutoff=cutoff_date_obj)
    
    return year_df