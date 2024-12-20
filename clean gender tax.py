import pandas as pd

df = pd.DataFrame({ 
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

with open("D:/งาน/SET CODE/gender tax.txt","r",encoding="utf-8") as f:
    lines = f.readlines()

count = 0
dateCount = 2

for line in lines:
    print(count)
    if(dateCount)%3 == 0:
        add["date"] = [line.strip()]
        dateCount += 1
        continue
    elif(count%5 ==0):
        df = pd.concat([df,add],ignore_index=True)
        add["men%"] = [line.strip()]
    elif(count%5 ==1):
        add["other"] = [line.strip()]
    elif(count%5 ==2):
        add["other%"] = [line.strip()]
    elif(count%5 ==3):
        add["total"] = [line.strip()]
    elif(count%5 ==4):
        add["total%"] = [line.strip()]
        if(dateCount%3 != 0):
            dateCount += 1
    count += 1

df = pd.concat([df,add],ignore_index=True)
print(df)

with pd.ExcelWriter("D:/งาน/SET/tax_by_gender01.xlsx") as writer:
    df.to_excel(writer,index=False)

