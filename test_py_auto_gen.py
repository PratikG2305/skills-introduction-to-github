# %%
import pandas as pd
import os


folder_path = r'C:\Users\p.gaikwad\Desktop\115'


os.remove('C:/Users/p.gaikwad/Desktop/115/output.xlsx')


dfs = []


for filename in os.listdir(folder_path):
    if filename.endswith('.xlsx') or filename.endswith('.xls'):  
        file_path = os.path.join(folder_path, filename)  
        df = pd.read_excel(file_path,engine='openpyxl')  
        dfs.append(df)  

combined_df = pd.concat(dfs, ignore_index=True)



# %%
data = pd.DataFrame()

cols = df.columns

data_dict = {col: [] for col in cols}

# %%
for col in df.columns :
    t = df[col].apply(lambda x: x.split('|') if isinstance(x, str) else x)
    st = t.apply(lambda x: [item.strip() for item in x] if isinstance(x, list) else x)
    st = st.apply(lambda x: [item.replace('One Of: ', '') for item in x] if isinstance(x, list) else x)

    data_dict[col] = st[0]


data_dict['publisher_id'] = [data_dict['publisher_id']]
data_dict['campaign_id'] = [data_dict['campaign_id']]
#data_dict

# %%
#[1, 2, 3]+(['-']*10)

# %%
# Check length of all lists
lengths = set(len(value) for value in data_dict.values() if isinstance(value, (list, tuple, str)))

lengths

# Add padding where necessary
for key in data_dict:
    if isinstance(data_dict[key], list):  # Only apply if value is a list
        if len(data_dict[key]) != max(lengths):
            data_dict[key] = data_dict[key] + ['NaN'] * (max(lengths) - len(data_dict[key]))

# Check if all lengths are now equal
set(len(value) for value in data_dict.values() if isinstance(value, (list, tuple, str)))

# %%
data = pd.DataFrame(data_dict)
data.replace("NaN","",inplace= True)

#data.fillna(0, inplace=True)

#data.to_excel('Lead Templet',sheet_name='Lead Templet',index=False)
#('Lead.xlsx',sheet_name= 'Lead',index= False)      

# %%


# %%
# Assuming data_dict contains your data
ab = pd.DataFrame(data_dict)

# Define the file path where you want to save the Excel file
excel_file_path = r'C:\Users\p.gaikwad\Desktop\115\output.xlsx'

# Save the DataFrame to an Excel file
data.to_excel(excel_file_path, index=False)


# %%



