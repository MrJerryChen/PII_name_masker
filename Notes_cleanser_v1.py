import pandas as pd

# Parameters
fileName = "Analysis_Latest_Test"
sheetName = "Sheet2"
colName = "Note"

# Read in excel file and first/last name databases
df = pd.read_excel(fileName+".xlsx",sheet_name=sheetName)
firstNames = pd.read_csv("FirstNameDatabase.csv")
lastNames = pd.read_csv("LastNameDatabase.csv")

# Iterate over each word in each row and match against databases
for idx,row in df.iterrows():
    tokens = row[colName].split()
    for i in range(len(tokens)):
        if tokens[i].lower().capitalize() in firstNames.values:
            tokens[i] = "Joe"
        elif tokens[i] in lastNames.values:
            tokens[i] = "Bloggs"
    df.at[idx,colName] = " ".join(tokens)
    print(str(idx + 1) + " completed.")

# Write dataframe to csv
df.to_csv(fileName+"_cleansed.csv",index=False)