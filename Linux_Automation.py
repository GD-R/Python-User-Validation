import os
import pandas as pd
import xlsxwriter

# Setting Working Directory
# os.chdir(r'Raw_Linux_Files')

# Reading Raw Files from the Folder
folder = os.chdir(r'Raw_Linux_Files')
files = os.listdir(folder)


print(files)
print(len(files))

# Setting Directory for Creating New Excel file
writer = pd.ExcelWriter(r'F:/Python-UV-Automation/MITMB/MITMB_Linux/MITMB_Linux_Validation.xlsx', engine='xlsxwriter')


# Reading all the Raw Linux Files and saving to one Workbook
li = []
for i in range(len(files)):
    user = pd.read_csv(files[i], encoding='unicode_escape', sep=':', header=0)
    user.insert(7, 'Server', files[i][:-7])
    user.to_excel(writer, sheet_name=files[i][:-7], index=False)

# Saving the Workbook
writer.save()


# Combining all the sheets from Workbook
combine = pd.concat(pd.read_excel(writer, sheet_name=None, engine='openpyxl'), ignore_index=True)

writer = pd.ExcelWriter(r'F:/Python-UV-Automation/MITMB/MITMB_Linux/MITMB_Linux_Validation.xlsx', engine='xlsxwriter')

# Deleting and Renaming Columns
combine.drop(combine.columns[[1, 2, 3, 5]], axis=1, inplace=True)
combine.columns = ['USERID', 'USERNAME', 'User_Details', 'SERVER']
# Removing Empty Rows
combine = combine.dropna()


# Assigning combine dataframe to new DataFrame
df_User_List = combine

# Filtering Service and System Account
searchfor = ['/sbin/nologin', '/sbin/shutdown', '/sbin/halt', '/usr/local/sbin/scponlyc', '/bin/sync', 'scv_', '/bin/false']
df_User_List = df_User_List[~df_User_List.User_Details.isin(searchfor) & ~df_User_List.USERID.str.contains(r'[_]')]
df_User_List = df_User_List[~df_User_List.USERID.str.contains(r'admin|auto|oracle|agent') & ~df_User_List.USERNAME.str.contains(r'[â]')]


# Deleting User_Details and Server Columns and Duplicates
df_User_List = df_User_List.drop(df_User_List.columns[[2, 3]], axis=1)
df_User_List = df_User_List.drop_duplicates(subset=['USERID'])
df_User_List.USERNAME = df_User_List.USERNAME.str.lstrip()


# Formatting Username
for i in range(len(df_User_List.index)):
    if (df_User_List.iloc[i, 1][:5].lower() == 'psmag') or (df_User_List.iloc[i, 1][:3].lower() == 'rfc'):
        df_User_List.iloc[i, 1] = (" ".join(df_User_List.iloc[i, 1].split()[2:4]))
    else:
        df_User_List.iloc[i, 1] = (" ".join(df_User_List.iloc[i, 1].split()[:2]))

df_User_List.USERNAME = df_User_List.USERNAME.str.replace(',', '')


# Reading previous Copy [Filename should match, Sheet name should be User_List ,  Col names should be USERID and TEAM]
old = pd.read_excel(r'F:/Python-UV-Automation/MITMB/MITMB_Linux/MITMB_Linux_Validation_Old.xlsx', sheet_name='User_List', usecols=['USERID', 'TEAM'], engine='openpyxl', header=0)


# VLookUp to identify the Team
result = pd.merge(df_User_List, old[['USERID', 'TEAM']], on='USERID', how='left')

# Saving Files
combine.to_excel(writer, sheet_name="Combined", index=False)
result.to_excel(writer, sheet_name="User_List", index=False)

writer.save()





