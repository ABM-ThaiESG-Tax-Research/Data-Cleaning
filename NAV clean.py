import os
import pandas as pd

# Folder containing the Excel files
new = pd.DataFrame({
    "Year": [],
    "name" : [],
    "Month": [],
    "Number of Funds-s": [],
    "Total Net Assets-s": [],
    "Mkt.Share-s":[],
    "Number of Funds-e": [],
    "Total Net Assets-e": [],
    "Mkt.Share-e":[],
    "Number of Funds-c": [],
    "Total Net Assets-c": [],
    "Mkt.Share-c":[],
})

new_rmf = pd.DataFrame({
    "Year": [],
    "name" : [],
    "Month": [],
  "Number of Funds-s": [],
  "Total Net Assets-s": [],
  "Mkt.Share-s":[],
  "Number of Funds-e": [],
  "Total Net Assets-e": [],
  "Mkt.Share-e":[],
  "Number of Funds-c": [],
  "Total Net Assets-c": [],
  "Mkt.Share-c":[],
})
new_ltf = pd.DataFrame({
    "Year": [],
    "name" : [],
    "Month": [],
  "Number of Funds-s": [],
  "Total Net Assets-s": [],
  "Mkt.Share-s":[],
  "Number of Funds-e": [],
  "Total Net Assets-e": [],
  "Mkt.Share-e":[],
  "Number of Funds-c": [],
  "Total Net Assets-c": [],
  "Mkt.Share-c":[],
})
for i in range (4,25):
    if(i<10):
        folder_path = "D:/งาน/SET/mf_200"+str(i)+"/"
    else:
        folder_path = "D:/งาน/SET/mf_20"+str(i)+"/"

    # folder_path = "D:/งาน/SET/mf_2009/"
    print("folder" +folder_path)
    for file_name in os.listdir(folder_path):
        rmf_al = False
        ltf_al = False
        if file_name.endswith(".xls")  or file_name.endswith(".xlsx"):  # Process only Excel files
            file_path = os.path.join(folder_path, file_name)
            if file_name == os.path.basename("D:/งาน/SET/mf_2009/data_2009.xlsx"):
                continue
            try:
                if(i < 9):
                    df = pd.read_excel(file_path, sheet_name='SUM by Classification')
                else:
                    df = pd.read_excel(file_path, sheet_name='Sum by Classification')
                
                # Iterate through the rows
                for index, row in df.iterrows():
                    month  = int(file_name[:2])
                    if row['Unnamed: 0'] == 'Retirement Mutual Fund (RMF)' and not rmf_al:  # Adjust based on file naming convention
                        print(month , '-', folder_path[-5:-1])
                        rmf_al = True
                        if i<9:
                            row_data = {
                                "Year" : '',
                                "name" : '',
                                "Month": month,
                                "Number of Funds-s": row['Unnamed: 1'],
                                "Total Net Assets-s": row['Unnamed: 2'],
                                "Mkt.Share-s": row['Unnamed: 3'],
                                "Number of Funds-e": 0,
                                "Total Net Assets-e": 0,
                                "Mkt.Share-e": 0,
                                "Number of Funds-c": 0,
                                "Total Net Assets-c": 0,
                                "Mkt.Share-c": 0
                            }
                        elif i <= 21:
                            if i ==21 and month > 2:
                                row_data = {
                                "Year" : '',
                                "name" : '',
                                "Month": month,
                                "Number of Funds-s": row['Unnamed: 2'],
                                "Total Net Assets-s": row['Unnamed: 3'],
                                "Mkt.Share-s": row['Unnamed: 4'],
                                "Number of Funds-e": row['Unnamed: 11'],
                                "Total Net Assets-e": row['Unnamed: 12'],
                                "Mkt.Share-e": row['Unnamed: 13'],
                                "Number of Funds-c": row['Unnamed: 14'],
                                "Total Net Assets-c": row['Unnamed: 15'],
                                "Mkt.Share-c": row['Unnamed: 16']
                            }
                            else:
                                row_data = {
                                    "Year" : '',
                                    "name" : '',
                                    "Month": month,
                                    "Number of Funds-s": row['Unnamed: 1'],
                                    "Total Net Assets-s": row['Unnamed: 2'],
                                    "Mkt.Share-s": row['Unnamed: 3'],
                                    "Number of Funds-e": row['Unnamed: 10'],
                                    "Total Net Assets-e": row['Unnamed: 11'],
                                    "Mkt.Share-e": row['Unnamed: 12'],
                                    "Number of Funds-c": row['Unnamed: 13'],
                                    "Total Net Assets-c": row['Unnamed: 14'],
                                    "Mkt.Share-c": row['Unnamed: 15']
                                }
                        else:
                            row_data = {
                                "Year" : '',
                                "name" : '',
                                "Month": month,
                                "Number of Funds-s": row['Unnamed: 2'],
                                "Total Net Assets-s": row['Unnamed: 3'],
                                "Mkt.Share-s": row['Unnamed: 4'],
                                "Number of Funds-e": row['Unnamed: 11'],
                                "Total Net Assets-e": row['Unnamed: 12'],
                                "Mkt.Share-e": row['Unnamed: 13'],
                                "Number of Funds-c": row['Unnamed: 14'],
                                "Total Net Assets-c": row['Unnamed: 15'],
                                "Mkt.Share-c": row['Unnamed: 16']
                            }
                        if(month == 1):
                            row_data["Year"] = folder_path[-5:-1]
                            row_data["name"] = "RMF"
                        new_rmf = pd.concat([new_rmf, pd.DataFrame([row_data])], ignore_index=True)
                    
                    if (row['Unnamed: 0'] == 'Long-Term Equity Fund (LTF)' or row["Unnamed: 0"] == 'Long Term Equity Fund (LTF)')and not ltf_al:  
                        print('ltf' ,month , '-', folder_path[-5:-1])
                        ltf_al = True
                        if i<9:
                            row_data = {
                                "Year" : '',
                                "name" : '',
                                "Month": month,
                                "Number of Funds-s": row['Unnamed: 1'],
                                "Total Net Assets-s": row['Unnamed: 2'],
                                "Mkt.Share-s": row['Unnamed: 3'],
                                "Number of Funds-e": 0,
                                "Total Net Assets-e": 0,
                                "Mkt.Share-e": 0,
                                "Number of Funds-c": 0,
                                "Total Net Assets-c": 0,
                                "Mkt.Share-c": 0
                            }
                        else:
                            row_data = {
                                "Year" : '',
                                "name" : '',
                                "Month": month,
                                "Number of Funds-s": row['Unnamed: 1'],
                                "Total Net Assets-s": row['Unnamed: 2'],
                                "Mkt.Share-s": row['Unnamed: 3'],
                                "Number of Funds-e": row['Unnamed: 10'],
                                "Total Net Assets-e": row['Unnamed: 11'],
                                "Mkt.Share-e": row['Unnamed: 12'],
                                "Number of Funds-c": row['Unnamed: 13'],
                                "Total Net Assets-c": row['Unnamed: 14'],
                                "Mkt.Share-c": row['Unnamed: 15']
                            }
                        if(month == 1):
                            row_data["Year"] = folder_path[-5:-1]
                            row_data["name"] = "LTF"
                        # Append to DataFrame
                        new_ltf = pd.concat([new_ltf, pd.DataFrame([row_data])], ignore_index=True)

#error on 06-2024 Sum by classifiaction => manually edit
                    if(rmf_al and ltf_al):
                        break
            except Exception as e:
                print(f"Error processing {file_name}: {e}")

    new = pd.concat([new_rmf,new_ltf],axis=0,ignore_index=True)

with pd.ExcelWriter("D:/งาน/SET/all_data.xlsx") as writer:
    new.to_excel(writer,index=False)