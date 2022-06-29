import os
import pandas as pd

# Setting Working Directory


# Reading Raw Files from the Folder
folder = os.chdir(r"Raw_Windows_Files")
files = os.listdir(folder)

print(files)
print(len(files))

# Setting Path for Creating New Excel file
writer = pd.ExcelWriter(r'F:/Python-UV-Automation/MITMB/MITMB_Windows/MITMB_Windows_Validation.xlsx', engine='xlsxwriter')

# Reading all the Windows Files into one Workbook
li = []
for i in range(len(files)):
    df = pd.read_csv(files[i], encoding='unicode_escape', sep='\t', header=0, usecols=['mail', 'samaccountname', 'name', 'description'])
    df.insert(0, "Domain", files[i][:-17])
    df.columns = ['AD Group', 'mail', 'USERID', 'name', 'description']
    li.append(df)
    df.to_excel(writer, sheet_name=(files[i].replace('-', "")[:-17]), index=False)

# Saving the File
writer.save()

# Combining all the Files from Workbook
combine = pd.concat(pd.read_excel(writer, sheet_name=None, engine='openpyxl'), ignore_index=True)

writer = pd.ExcelWriter(r'F:/Python-UV-Automation/MITMB/MITMB_Windows/MITMB_Windows_Validation.xlsx', engine='xlsxwriter')

# *************************************5/20/2022*********************************************

# ['AD Group', 'mail', 'USERID', 'name', 'description']

combine['mail'] = combine['mail'].astype('str').str.replace(r"b'", r"", regex=False)
combine['USERID'] = combine['USERID'].astype('str').str.replace(r"b'", r"", regex=False)
combine['name'] = combine['name'].astype('str').str.replace(r"b'", r"", regex=False)
combine['name'] = combine['name'].astype('str').str.replace(r".", r"", regex=False)
combine['name'] = combine["name"].str.replace('\s+', ' ', regex=True)

combine['description'] = combine['description'].astype('str').str.replace(r"b'", r"", regex=False)

combine = combine.replace("'", "", regex=True)

print(combine.mail)

# *************************************************************************************

# Reading previous Copy [Filename should match, Sheetname should be User_List ,  Col names should be USERID and TEAM]
old = pd.read_excel(r'F:/Python-UV-Automation/MITMB/MITMB_Windows/MITMB_Windows_Validation_Old.xlsx', sheet_name='User_List', usecols=['USERID', 'TEAM'], engine='openpyxl', header=0)
old = old.drop_duplicates(subset=['USERID'])


# Vlookup to identify the Team
result = pd.merge(combine, old[['USERID', 'TEAM']], on='USERID', how='left')
result.sort_values(by='TEAM', inplace=True)


# Saving Files
result.to_excel(writer, sheet_name="User_List", index=False, header=True)
writer.save()

