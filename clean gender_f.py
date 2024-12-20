import pandas as pd

df = pd.DataFrame({ 
    "year" : [],
    "type":[],
    "woman":[],
    "woman%":[],
    "men":[],
    "men%":[],
    "other":[],
    "other%":[],
    "total":[],
    "total%":[],
    "date":[]
})

add = pd.DataFrame({ 
     "year" : [],
    "type":[],
    "woman":[],
    "woman%":[],
    "men":[],
    "men%":[],
    "other":[],
    "other%":[],
    "total":[],
    "total%":[],
    "date":[]
})

with open("D:/งาน/SET Code/gender tax_f.txt","r",encoding="utf-8") as f:
    lines = f.readlines()

count = 0
dateCount = 2

for line in lines:
    print(count)
    if(dateCount)%3 == 0:
        add["year"] = [line.strip()]
        dateCount += 1
        continue
    elif(count%4 ==0):
        df = pd.concat([df,add],ignore_index=True)
        add["type"] = [line.strip()]
    elif(count%4 ==1):
        add["woman"] = [line.strip()]
    elif(count%4 ==2):
        add["woman%"] = [line.strip()]
    elif(count%4 ==3):
        add["men"] = [line.strip()]
        if(dateCount%3 != 0):
            dateCount += 1
    count += 1

df = pd.concat([df,add],ignore_index=True)
print(df)

with pd.ExcelWriter("D:/งาน/SET/tax_by_gender02.xlsx") as writer:
    df.to_excel(writer,index=False)

