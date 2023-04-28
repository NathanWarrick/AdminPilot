from xls2xlsx import XLS2XLSX
import os
import pandas as pd #for dataframe operations and such.



# Determine the newest file added to the directory
files = os.listdir(r"workspace\bots\qkr_bot\Downloads")
paths = [os.path.join(r"workspace\bots\qkr_bot\Downloads", basename) for basename in files]
name = max(paths, key=os.path.getctime)


# Rename the file 
excursionname = name[32:name.find('-')] # Trim string
print(excursionname)
name_new = r"workspace\bots\qkr_bot\Downloads" + excursionname + '.xls'
print(name_new)
xlsxfile = r"workspace\bots\qkr_bot\Downloads" + excursionname + '.xlsx'
print(xlsxfile)


# Copy/Move the file to the downloads directory, only uncomment one of these
os.replace(name, name_new)

# Convert the file to xlsx and remove the xls version
XLS2XLSX(name_new).to_xlsx(xlsxfile)
os.remove(name_new)

# Split students name into first and last. Added an exception for local excursions that have a different format

qkr_df = pd.read_excel(xlsxfile) 

if excursionname == '\LocalExcursionForm':
    print("Local Excursion Form Detected")
    qkr_df[['First Name','Last Name']] = qkr_df['Student Name:'].loc[qkr_df['Student Name:'].str.split().str.len() == 2].str.split(expand=True) # Split students name into first and last
    qkr_df['First Name'].fillna(qkr_df['Student Name:'],inplace=True)
    qkr_df = qkr_df[["First Name", "Last Name", "Parent/Carer's Name:" ,"Phone Number 1:", "Name:", "Relationship to student:", "Phone Number:"]]
    qkr_df = qkr_df.rename(columns={"Parent/Carer's Name:": "Guardian's Name",
                            "Parent/Carer's business hours number:": "Contact Number",
                            "Name:": "Emergency Contact Name",
                            "Relationship to student:": "Relationship",
                            "Phone Number:": "Contact Number"})

else:
    qkr_df[['First Name','Last Name']] = qkr_df['Students Full Name:'].loc[qkr_df['Students Full Name:'].str.split().str.len() == 2].str.split(expand=True) # Split students name into first and last
    qkr_df['First Name'].fillna(qkr_df['Students Full Name:'],inplace=True)


    qkr_df = qkr_df[["First Name", "Last Name", "Parent/Carer's Full Name:" ,"Parent/Carer's business hours number:"]]
    qkr_df = qkr_df.rename(columns={"Parent/Carer's Full Name:": "Guardian's Name",
                            "Parent/Carer's business hours number:": "Contact Number"})



# Compare master file in /Excursions to downloaded and formatted file, exception for if file does not exist
try:
    masterfile = 'Excursions' + excursionname + '.xlsx'
    master_df = pd.read_excel(masterfile)

    frames = [master_df, qkr_df]
    result = pd.concat(frames)
    result = result.drop_duplicates()
    result.to_excel('Excursions' + excursionname + '.xlsx', index=False)  
    
    
except:
    qkr_df.to_excel("Excursions" + excursionname + '.xlsx', index=False)  

os.remove(xlsxfile) 

