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


df = pd.read_excel(xlsxfile) 

# Split students name into first and last
df[['First Name','Last Name']] = df['Students Full Name:'].loc[df['Students Full Name:'].str.split().str.len() == 2].str.split(expand=True) # Split students name into first and last
#df['First Name'].fillna(df['Name'],inplace=True)


df = df[["First Name", "Last Name"]]
#df = df.rename(columns={"Students Full Name:": "Name"})


#df[['First Name','Last Name']] = df['Name'].loc[df['Name'].str.split().str.len() == 2].str.split(expand=True)
print(df)
df.to_excel(xlsxfile)  
