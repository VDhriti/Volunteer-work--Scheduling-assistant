import xlrd

loc = ('C:\\Users\\dhrit\\Desktop\\Availability from 15 SepTuesday.xlsx')
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

#print(sheet.cell_value(1, 0))   #(2, A/ ROW, COLUMN)

lastRow = sheet.nrows
names = []
days = []

for i in range(2, lastRow):
    if (sheet.cell_value(i, 1) != ''):
        names.append(sheet.cell_value(i, 1).lower())

    Value = sheet.cell_value(i, 2)
    if (Value != ''):
        if Value == 'All':
            days.append(['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'])

        elif Value == 'Weekdays':
            days.append(['mon', 'tue', 'wed', 'thu', 'fri'])

        elif Value == 'Weekends':
            days.append(['sat', 'sun'])

        elif '+' in Value:
            tempList = Value.split('+').lower()
            days.append(tempList)

        else:
            days.append(Value.lower())
        
rowDict = {}
for i in range(len(names)):
    rowDict[names[i]] = days[i]

#print(rowDict)
