import xlrd

loc = ('.\\Test Excel.xlsx')
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

#print(sheet.cell_value(1, 0))   #(2, A/ ROW, COLUMN)

lastRow = sheet.nrows
names = []
days = []
mealTimes = []

for h in range(2, lastRow):
    if (sheet.cell_value(h, 1) != ''):
        names.append(sheet.cell_value(h, 1).lower())

for i in range(2, len(names)+2):
    Value = sheet.cell_value(i, 2).lower()
    if Value == 'all':
        days.append(['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'])
    elif Value == 'weekdays':
        days.append(['mon', 'tue', 'wed', 'thu', 'fri'])
    elif Value == 'weekends':
        days.append(['sat', 'sun'])
    elif '+' in Value:
        Value.replace(' ', '')
        tempList = Value.split('+')
        days.append(tempList)
    elif Value == '':
        days.append(['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'])
    else:
        days.append(Value)
   
    mealTimes.append(
        [
            sheet.cell_value(i, 3).lower(), 
            sheet.cell_value(i, 4).lower(), 
            sheet.cell_value(i, 5).lower()
        ]
    )   
    

rowDict = {}
for j in range(len(names)):
    rowDict[j] = [names[j], days[j], mealTimes[j]]


startDay = input("Please enter the first three letters of the start day: ").lower()
daysOfTheWeek = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun']

dayIndex = daysOfTheWeek.index(startDay)               # monday = 0, tuesday = 1 and so on

dayCounter = 0
bSpec, bAll, lSpec, lAll, dSpec, dAll = [], [], [], [], [], []

while dayCounter < 10:
    b, l, d = [bSpec, bAll], [lSpec, lAll], [dSpec, dAll]
    avMeals = [b, l, d]
    
    for k in range(len(names)):
        if daysOfTheWeek[dayIndex] in rowDict[k][1]:
            #pplAvailableThisDay.append(rowDict[k][0]
            for l in range(3):
                if rowDict[k][2][l] == 'y':
                    if rowDict[k][1] == ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun']:
                        avMeals[l][1].append(rowDict[k][0])
                    else:
                        avMeals[l][0].append(rowDict[k][0])
             
    print('Day', dayCounter+1, ':', daysOfTheWeek[dayIndex].capitalize())
    print('   ', 'Breakfast: ')
    print('   ', '   ', 'Specific Availability: ', bSpec)
    print('   ', '   ', 'Free for All', bAll)
    print('')

    print('   ', 'Lunch: ')
    print('   ', '   ', 'Specific Availability: ', lSpec)
    print('   ', '   ', 'Free for All', lAll)
    print('')

    print('   ', 'Dinner: ')
    print('   ', '   ', 'Specific Availability: ', dSpec)
    print('   ', '   ', 'Free for All', dAll)
    print('')


    bSpec, bAll, lSpec, lAll, dSpec, dAll = [], [], [], [], [], []

    print('')
    print('-------------------------------------')
    print('')

    dayIndex = (dayIndex + 1)%7
    dayCounter+=1
