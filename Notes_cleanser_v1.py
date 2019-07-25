import pandas as pd
import sys
import random

# Parameters
fileName = sys.argv[1]
sheetName = "Sheet1"
colName = "Note"

# Read in excel file and first/last name databases
df = pd.read_excel(fileName,sheet_name=sheetName)
firstNames = pd.read_csv("FirstNameDatabase.csv")
lastNames = pd.read_csv("LastNameDatabase.csv")
names = ["Apple","Apricot","Avocado","Banana","Blueberry","Berry","Cantaloupe","Cherry","Coconut","Date","Grape","Guava","Kiwi","Lemon","Lime","Mango","Melon","Nectarine","Orange","Papaya","Peach","Pear","Pineapple","Plum","Raisin","Raspberry","Strawberry","Tangerine"]

# Iterate over each word in each row and match against databases
for idx,row in df.iterrows():
    tokens = row[colName].split()
    for i in range(len(tokens)):
        if tokens[i].lower().capitalize() in firstNames.values:
            tokens[i] = random.choice(names)
        elif tokens[i].lower().capitalize() in lastNames.values:
            tokens[i] = random.choice(names)
    df.at[idx,colName] = " ".join(tokens)
    print(str(idx + 1) + " completed.")

# Write dataframe to csv
df.to_excel("output.xlsx",index=False)
# ("output.csv",index=False)