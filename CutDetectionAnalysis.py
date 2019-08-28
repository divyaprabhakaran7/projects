import os

# Retrieve current working directory
cwd = os.getcwd()
cwd

# Store foldername and extension
folderLoc = '/Desktop/CutData'

# Change directory to folder path
os.chdir(cwd + folderLoc)

# List all files and directories in current directory
os.listdir('.')

# Import xlrd, xlsx writer modules
import xlrd
import xlsxwriter

# Give location of Excel file
loc = ('Frames and Cut Types.xlsx')

# Open workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# Workbook takes the filename we want to create
workbook = xlsxwriter.Workbook('DataAnalysis.xlsx')

# Workbook object adds a new worksheet
worksheet = workbook.add_worksheet()

# Use worksheet object to write data intot the new Excel spreadsheet
worksheet.write('A1', 'v1v2')
worksheet.write('B1', 'Cut Type')
worksheet.write('C1', 'Pre Mean Diff')
worksheet.write('D1', 'Post Mean Diff')
worksheet.write('E1', 'Pre Mean Flow')
worksheet.write('F1', 'Post Mean Flow')
worksheet.write('G1', 'v2v1')
worksheet.write('H1', 'Cut Type')
worksheet.write('I1', 'Pre Mean Diff')
worksheet.write('J1', 'Post Mean Diff')
worksheet.write('K1', 'Pre Mean Flow')
worksheet.write('L1', 'Post Mean Flow')

# Store row and column position
row = 1
col = 0

# Import mean function from statistics
from statistics import mean

# Open v1v2 data
v1v2wb = xlrd.open_workbook('v1v2_all.xlsx')
v1v2sheet = v1v2wb.sheet_by_index(0)

for i in range(sheet.nrows):
    counter = sheet.cell_value(row, col)
    compList = []
    # Append cut frame and type
    compList.append(int(counter))
    compList.append(sheet.cell_value(row, col + 1))
    rf = []
    of = []
    rd = []
    od = []
    # Calculate mean prediff
    rdcells = v1v2sheet.col_slice(1, int(counter - 12), int(counter))
    for cell in rdcells:
        rd.append(cell.value)
    compList.append(mean(rd))
    # Calculate mean postdiff
    odcells = v1v2sheet.col_slice(1, int(counter), int(counter + 12))
    for cell in odcells:
        od.append(cell.value)
    compList.append(mean(od))
    # Calculate mean preflow
    rfcells = v1v2sheet.col_slice(2, int(counter - 13), int(counter - 1))
    for cell in rfcells:
        rf.append(cell.value)
    compList.append(mean(rf))
    # Calculate postFlow
    ofcells = v1v2sheet.col_slice(2, int(counter + 1), int(counter + 13))
    for cell in ofcells:
        of.append(cell.value)
    compList.append(mean(of))
    # Write to new Excel file
    worksheet.write_row(row, col, compList)
    row += 1

# Open v2v1 data
v2v1wb = xlrd.open_workbook('v1v2_all.xlsx')
v2v1sheet = v2v1wb.sheet_by_index(0)

col = 2
row = 1

for i in range(sheet.nrows):
    counter = sheet.cell_value(row, col)
    compList = []
    compList.append(int(counter))
    compList.append(sheet.cell_value(row, col + 1))
    rf = []
    of = []
    rd = []
    od = []

    rdcells = v2v1sheet.col_slice(1, int(counter - 12), int(counter))
    for cell in rdcells:
        rd.append(cell.value)
    compList.append(mean(rd))

    odcells = v2v1sheet.col_slice(1, int(counter), int(counter + 12))
    for cell in odcells:
        od.append(cell.value)
    compList.append(mean(od))

    rfcells = v2v1sheet.col_slice(2, int(counter - 13), int(counter - 1))
    for cell in rfcells:
        rf.append(cell.value)
    compList.append(mean(rf))

    ofcells = v2v1sheet.col_slice(2, int(counter + 1), int(counter + 13))
    for cell in ofcells:
        of.append(cell.value)
    compList.append(mean(of))

    worksheet.write_row(row, col + 4, compList)
    row += 1

# Close the Excel file
workbook.close()
