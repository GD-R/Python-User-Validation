import os
import pandas as pd


folder = os.chdir(r'Raw_DBA_Files')
files = os.listdir(folder)

old = pd.read_excel(r'F:/Python-UV-Automation/MITMB/MITMB_Database/MITMB_Database_Validation_Old.xlsx', sheet_name='User_List',
                    usecols=['USERID', 'TEAM', 'USERNAME'], engine='openpyxl', header=0)

# ******************************

old = old.fillna('')
old = old.replace("\t", "", regex=True)
old['USERID'] = old['USERID'].str.rstrip()
# old = old.replace("  ", "", regex=True)

# ******************************


col_names = ["USERID", "col2", "ROLES", ...]

writer = pd.ExcelWriter(r'F:/Python-UV-Automation/MITMB/MITMB_Database/MITMB_Database_Validation.xlsx', engine='xlsxwriter',
                        engine_kwargs={'options': {'strings_to_numbers': False, 'strings_to_formulas': False,
                                                   'strings_to_urls': False}})

for i in range(len(files)):
    file = pd.read_csv(files[i], encoding='unicode_escape', names=col_names)
    if file.columns.tolist()[3] == Ellipsis:
        file.drop(file.columns[3], axis=1, inplace=True)
    file = file.drop(columns="col2")
    file = file.fillna('')
    file = file.replace("\t", "", regex=True)
    file = file.replace(" ", "", regex=True)
    file.to_excel(writer, sheet_name="Sheet" + str(i), index=False)

writer.save()

combine_PD_NPD = pd.concat(pd.read_excel(writer, sheet_name=None, engine='openpyxl'), ignore_index=True)
print(combine_PD_NPD.tail(10))


li = []
temp = 4

for i in range(len(combine_PD_NPD.index)):
    if "rowsselected" in str(combine_PD_NPD.iloc[i, 0]):
        li.append(i)

print(li)

writer = pd.ExcelWriter(r'F:/Python-UV-Automation/MITMB/MITMB_Database/MITMB_Database_Validation.xlsx', engine='xlsxwriter',
                        engine_kwargs={'options': {'strings_to_numbers': False, 'strings_to_formulas': False,
                                                   'strings_to_urls': False}})


for i in li:
    current = combine_PD_NPD.iloc[temp:i]
    current.reset_index(drop=True, inplace=True)
    current = current.replace("\t", "", regex=True)
    current = current.replace(" ", "", regex=True)
    current = current.fillna('')
    current.insert(2, 'DB', current.loc[0, 'USERID'].upper())
    current.drop([0, 1, 2], inplace=True)
    current.reset_index(drop=True, inplace=True)
    current.to_excel(writer, sheet_name=current.loc[0, 'DB'], index=False)
    temp = i + 6

writer.save()

combine = pd.concat(pd.read_excel(writer, sheet_name=None, engine='openpyxl'), ignore_index=True)
combine = combine.dropna()

combine = pd.merge(combine, old[['USERID', 'USERNAME', 'TEAM']], on='USERID', how='left')
combine.sort_values(by=['TEAM'], inplace=True)
combine.reset_index(drop=True, inplace=True)

with pd.ExcelWriter(r'F:/Python-UV-Automation/MITMB/MITMB_Database/MITMB_Database_Validation.xlsx', mode="a", engine="openpyxl") as writer:
    combine.to_excel(writer, sheet_name="combine", index=False)

df_User_List = combine

df_User_List = df_User_List[~df_User_List.ROLES.str.contains(r'[_]')]

df_User_List = df_User_List[~df_User_List.USERID.str.contains(r'OWN|READ1|WORK|SYS|INTF|_|XDB|USER|DATA')]

# df_User_List = pd.merge(df_User_List, old[['USERID', 'USERNAME', 'TEAM']], on='USERID', how='left')
df_User_List.sort_values(by=['USERNAME'], inplace=True)
df_User_List = df_User_List.drop(df_User_List.columns[[1, 2]], axis=1)
df_User_List = df_User_List.drop_duplicates(subset=['USERID'])
df_User_List.reset_index(drop=True, inplace=True)

print(df_User_List.head(10))


with pd.ExcelWriter(r'F:/Python-UV-Automation/MITMB/MITMB_Database/MITMB_Database_Validation.xlsx', mode="a", engine="openpyxl") as writer:
    df_User_List.to_excel(writer, sheet_name="User_List", index=False)

