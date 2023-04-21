from xls2xlsx import XLS2XLSX
import os
import shutil
import pandas as pd #for dataframe operations and such.

os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Determine the newest file added to the directory
files = os.listdir('Downloads\Waiting')
paths = [os.path.join('Downloads\Waiting', basename) for basename in files]
name = max(paths, key=os.path.getctime)

# Rename the file 
name_trimmed = name[17:name.find('-')] # Trim string
name_new = 'Downloads' + name_trimmed + '.xls'
xlsxfile = 'Downloads' + name_trimmed + '.xlsx'


# Copy/Move the file to the downloads directory, only uncomment one of these
shutil.copyfile(name, name_new)
#os.replace(name, name_new)

# Convert the file to xlsx and remove the xls version
XLS2XLSX(name_new).to_xlsx(xlsxfile)
os.remove(name_new)


qkr_df = pd.read_excel(xlsxfile) 

# Split students name into first and last. Added an exception for local excursions that have a different format
try:
    qkr_df[['First Name','Last Name']] = qkr_df['Students Full Name:'].loc[qkr_df['Students Full Name:'].str.split().str.len() == 2].str.split(expand=True) # Split students name into first and last
    qkr_df['First Name'].fillna(qkr_df['Students Full Name:'],inplace=True)
except: # Change this later to explicitly work for local excursion forms
    qkr_df[['First Name','Last Name']] = qkr_df['Student Name:'].loc[qkr_df['Student Name:'].str.split().str.len() == 2].str.split(expand=True) # Split students name into first and last
    qkr_df['First Name'].fillna(qkr_df['Student Name:'],inplace=True)

try:
    qkr_df = qkr_df[["First Name", "Last Name", "Parent/Carer's Full Name:" ,"Parent/Carer's business hours number:"]]
    qkr_df = qkr_df.rename(columns={"Parent/Carer's Full Name:": "Guardian's Name",
                            "Parent/Carer's business hours number:": "Contact Number"})
except:
    qkr_df = qkr_df[["First Name", "Last Name", "Parent/Carer's Name:" ,"Phone Number 1:", "Name:", "Relationship to student:", "Phone Number:"]]
    qkr_df = qkr_df.rename(columns={"Parent/Carer's Name:": "Guardian's Name",
                            "Parent/Carer's business hours number:": "Contact Number",
                            "Name:": "Emergency Contact Name",
                            "Relationship to student:": "Relationship",
                            "Phone Number:": "Contact Number"})



try:

    masterfile = 'Excursions' + name_trimmed + '.xlsx'
    master_df = pd.read_excel(masterfile)

    frames = [master_df, qkr_df]
    result = pd.concat(frames)
    result = result.drop_duplicates()
    result.to_excel('Excursions' + name_trimmed + '.xlsx', index=False)  
    
    
except:
    qkr_df.to_excel('Excursions' + name_trimmed + '.xlsx', index=False)  

os.remove(xlsxfile) 

