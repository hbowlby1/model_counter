import openpyxl
import pandas as pd

insight_items = []
ePIC_items = []
isInInsightAndEpic = []


def insightReader(csvFile):
    csvFile = pd.read_csv(csvFile)
    
    for index, row in csvFile.iterrows():
        insight_items.append(row['Item Number'])
    # return csvFile
    return insight_items

def ePICReader(excelFile):
    xcel1 = pd.read_excel(excelFile, sheet_name=0, usecols=["Material"])
    xcel2 = pd.read_excel(excelFile, sheet_name=1, usecols=["Summary Stock Listing"])

    # iterate through the first sheet and append data to array
    for index, row in xcel1.iterrows():
        ePIC_items.append(row["Material"])

    for index, row in xcel2.iterrows():
        ePIC_items.append(row["Summary Stock Listing"])

    # return ePIC_items
    return(ePIC_items)


insightData = insightReader("spreadsheets/items_from_insight.csv")
epicData = ePICReader("spreadsheets\sheet_goods.xlsx")

def compare():
    for i in insightData:
        for j in epicData:
            if(i == j):
                isInInsightAndEpic.append(j)
            else:
                pass
    return isInInsightAndEpic

df = pd.DataFrame(compare())
df.to_excel(excel_writer="sheet_goods_in_insight.xlsx")

