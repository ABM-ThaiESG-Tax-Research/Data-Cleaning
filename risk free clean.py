import pandas as pd

df = pd.read_excel("Data Cleaning/TableMPC_2567_6.xlsx")
months = {"Jan": 1 , "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6, "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}

add = pd.DataFrame({
    "Date": [],
    "Policy rate":[],
})

new = pd.DataFrame({
    "Date": [],
    "Policy rate":[],
})

start_year = 2000
start_month = 1

for (index, row) in df.iterrows():
    if pd.isna(row["Date"]): continue
    print(len(str(row["Date"])))
    if(len(str(row["Date"])) == 19):
        year = str(row['Date'])[:4]
        month = str(row['Date'])[5:7]
    else:
        year = str(row['Date'])[-3:]
        if(str(row['Date'])[2] == "-"):
            month = months[str(row['Date'])[3:6]]
        else:
            month = months[str(row['Date'])[2:5]]
    print("year" , year,"\n month" ,month,"\n full" , row["Date"])
    if(int(year) != start_year):
        add["Date"] = [str(i) + "/" + str(start_year) for i in range(start_month,13)]
        add["Policy rate"] = [row["Policy Rate"] for i in range(start_month,13)]
        print(add)
        start_year = int(year)
        new = pd.concat([new, add])
        add = pd.DataFrame({
            "Date": [],
            "Policy rate":[],
        })
    if int(month) != start_month:
        add["Date"] = [str(i) + "/" + year for i in range(start_month, int(month))]
        add["Policy rate"] = [row["Policy Rate"] for i in range(start_month, int(month))]
        print(add)
        start_month = int(month)
        new = pd.concat([new, add])
        add = pd.DataFrame({
            "Date": [],
            "Policy rate":[],
        })
    
with pd.ExcelWriter("Data Cleaning/TableMPC_2567_6_cleaned.xlsx") as writer:
    new.to_excel(writer, sheet_name='Sheet1', index=False)

print(new)