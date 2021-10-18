import pandas as pd
import os
import sys

# Read from Excel
xl = pd.ExcelFile("output.xls")

# Parsing Excel Sheet to DataFrame
dataFrameSheet = xl.parse(xl.sheet_names[0])

# finding the drop row list
dropRowNoList = []
for rowNo, row in dataFrameSheet.iterrows():
    isAllZero = True
    for colNo, col in enumerate(row):
        if colNo == 0:
            continue
        if col != 0:
            isAllZero = False

    if isAllZero:
        dropRowNo = rowNo + 2
        print("Adding rowNo " + str(rowNo) + " to dropRowNoList")
        dropRowNoList.append(rowNo)

# dropping row will all zeros
print(dropRowNoList)
dataFrameSheet = dataFrameSheet.drop(dropRowNoList)


currentDirectory = os.path.dirname(sys.argv[0])
outPutFileName = "optimized_output"
outputFilePath = os.path.join(currentDirectory, outPutFileName + '.xls')

# Updating the excel sheet with the updated DataFrame
dataFrameSheet.to_excel(outPutFileName + ".xls", sheet_name='Sheet1', index=False)

# Open it via the operating system (will only work on Windows)
os.startfile(outputFilePath)