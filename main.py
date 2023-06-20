import pandas as pd
import math

def excelReader(sheet):
    # Read the excel file and save each sheet
    spreadsheet = pd.ExcelFile(sheet)
    modelAndQuanityDict = {}

    for sheet_name in spreadsheet.sheet_names:
        # Read each sheet and save it as a dataframe
        df = pd.read_excel(spreadsheet, sheet_name=sheet_name)

        # Get the index for "Item Number":
        item_number_index = df.columns.get_loc("Item Number")

        for row_index, row in df.iterrows():
            # Get the model name from the "Item name" column
            model_name = row.iloc[item_number_index]

            # Check if the model name already exists in dictionary
            if model_name not in modelAndQuanityDict:
                # Add it and then set the quantity to 0
                modelAndQuanityDict[model_name] = {"total": 0}

            # Loop through the columns and find the counts
            for col_name, cell in row.items():
                # Check the column name that would contain count
                if (
                    col_name == "Available"
                    or col_name == "Trans Count Qty"
                    or col_name == "Trans Count"
                ):
                    if(math.isnan(cell)):
                        modelAndQuanityDict[model_name]["total"] += 0
                    else:
                        modelAndQuanityDict[model_name]["total"] += cell

    return modelAndQuanityDict

# convert the returned data to excel
totals = excelReader("spreadsheets/allCountsForProgramNoFirstPage.xlsx")

df = pd.DataFrame.from_dict(totals, orient="index")

df.to_excel("newTotals.xlsx", index_label="Item_number")
