
#import python libraries
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog
import xlwt 
from datetime import date 
from datetime import datetime


#Create a function to prompt the user to upload a file
def file_upload():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title = "Select File")
    return file_path

#Run the function to get the file path
file_name = file_upload()

#Read the file
df = pd.read_csv(file_name)



#Hide the specified columns
df.drop(columns = ['recommendationName','resourceId','recommendationId','azurePortalRecommendationLink'], axis = 1, inplace = True)

#Filter the column N for the value Unhealthy
df = df[df['state'] == 'Unhealthy']


#generate todays date
today = date.today() 

# Define a function to apply the formatting 
def highlight_cells(val):
    color = 'red' if val == 'High' else 'white'
    return 'background-color: %s; color: %s' % (color, 'white' if val == 'High' else 'black')

# Apply the function to the column of interest
df.style.applymap(highlight_cells, subset=['severity'])



# Save dataframe to excel file with today's date
today_date = datetime.now().strftime("%Y_%m_%d-%I_%M_%S_%p")
df.to_excel(f'Defender_Security_Recommendations_{today_date}.xlsx')

#show succesfull export of new excel file.
print('Data successfully filtered and saved to an excel file')

