# Import all the relevant modules
import pandas as pd
from datetime import datetime
from datetime import time
import numpy as np
import matplotlib.pyplot as plt
import os

# Define Folder
dir_path = r'overtime'

# Create a framework for the final table, which will have the desired data
table = pd.DataFrame(columns=['Pay No','Date From','Date To','Payroll Number','Position','Surname','First Name','Unit','Date of Overtime','Reason for Overtime','Unable to Handover Because','Start Time','End Time','Total Time'])

# Main looping sequence to iterate through the files and collate the data systematically
for files in os.listdir(dir_path):
    f = os.path.join(dir_path, files)
    otin = pd.read_excel(f)

    otin = otin.drop(otin.columns[[13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23]], axis = 1)

    # Contact data harvesting
    data_payroll = otin.iloc[2,0]
    data_position = otin.iloc[2,1]
    data_surname = otin.iloc[2,2]
    data_firstname = otin.iloc[2,3]
    data_unit = otin.iloc[2,4]

    otin.columns = otin.iloc[7]
    otin = otin.iloc[9:26]
    otin = otin.dropna(subset=['Date of Overtime'])

    # Looping through each overtime claim and combining it with the contact data - for each line on the final table
    i = 0
    for i in range(0, len(otin.index)):
        data_date = otin.iloc[i,0]
        data_reason = otin.iloc[i,1]
        data_handover = otin.iloc[i,2]

        data_start = otin.iloc[i,8]
        data_start = str(data_start)
        data_start = datetime.strptime(data_start, "%H:%M:%S")
        data_end = otin.iloc[i,11]
        data_end = str(data_end)
        data_end = datetime.strptime(data_end, "%H:%M:%S")
        data_timediff = data_end - data_start
        data_timediff = data_timediff.total_seconds()/3600

        data_start = data_start.time()
        data_start = data_start.strftime("%H:%M")
        data_end = data_end.time()
        data_end = data_end.strftime("%H:%M")

        data_payno = ""
        data_datefrom = ""
        data_dateto = ""

        # Create a new row of the gathered information above
        new_row = {'Pay No':data_payno,'Date From':data_datefrom,'Date To':data_dateto,'Payroll Number':data_payroll,'Position':data_position,'Surname':data_surname,'First Name':data_firstname,
                   'Unit':data_unit,'Date of Overtime':data_date,'Reason for Overtime':data_reason,'Unable to Handover Because':data_handover,
                   'Start Time':data_start,'End Time':data_end,'Total Time':data_timediff}

        # Append the rows to the existing table
        table = table.append(pd.DataFrame([new_row],index=['Index'],columns=table.columns))

        i = i + 1

    table = table.reset_index(drop=True)

# Export the final table as an Excel file
table.to_excel('Overtime Collated.xlsx')
