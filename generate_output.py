import os
import sys
import csv
import xlwt

currentDirectory = os.path.dirname(sys.argv[0])
outPutFileName = "output"
outputFilePath = os.path.join(currentDirectory, outPutFileName + '.xls')

# getting the files from the directory
textFilePath = os.path.join(currentDirectory, 'Txt')
textFileList = [f for f in os.listdir(textFilePath) if os.path.isfile(os.path.join(textFilePath, f))]

# Create a workbook object
workBook = xlwt.Workbook()

# Add a sheet object
workSheet = workBook.add_sheet(outPutFileName, cell_overwrite_ok=True)

# Writing Gene Expression in row-0 col-0
workSheet.write(0, 0, 'Gene Expression')

rowCount = 1
columnCount = 0
for index, textFile in enumerate(textFileList):
    print('Working on item: ' + str(index) + ' fileName: ' + textFile)
    workSheet.write(0, index + 1, os.path.splitext(textFile)[0])
    currentTextFilePath = os.path.join(textFilePath, textFile)
    # Get a CSV reader object set up for reading the input file with tab delimiters
    dataReader = csv.reader(open(currentTextFilePath, 'rt'), delimiter='\t', quotechar='"')

    emptyRow = 0
    # Process the file and output to Excel sheet
    for rowNo, row in enumerate(dataReader):
        for colNo, colItem in enumerate(row):
            # # unblock the following block if want to remove 0 from data
            # if colNo == 0 and float(row[1]) == 0:
            #     continue
            # if colNo == 1 and float(colItem) == 0:
            #     emptyRow += 1
            #     continue
            if columnCount != 0 and colNo == 0:
                continue
            workSheet.write(rowNo - emptyRow + rowCount, colNo + columnCount, colItem)

    columnCount += 1

# Write the output file.
workBook.save(outputFilePath)

# Open it via the operating system (will only work on Windows)
os.startfile(outputFilePath)

