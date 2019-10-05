import matplotlib.pyplot as plt
import os
import shelve


def draw_graph(data, labels, title):
    num_bars = len(data)
    positions = range(1, num_bars+1)
    plt.bar(positions, data, align='center')
    plt.xticks(positions, labels)
    plt.title(title)
    plt.xlabel('Months')
    plt.ylabel('Amount (R)')
    plt.grid()
    plt.show()


directory = 'C:\\AnniqueSpreadSheets\\data'
os.chdir(directory)
print(os.getcwd())
memNum = []
shelf = shelve.open('income')
memNum = memNum + (shelf.__getitem__('members'))
print(memNum)

memIncome1 = []
memIncome1 = memIncome1 + shelf.__getitem__('month1')
memIncome2 = []
memIncome2 = memIncome2 + shelf.__getitem__('month2')
memIncome3 = []
memIncome3 = memIncome3 + shelf.__getitem__('month3')
memIncome4 = []
memIncome4 = memIncome4 + shelf.__getitem__('month4')
memIncome5 = []
memIncome5 = memIncome5 + shelf.__getitem__('month5')
memIncome6 = []
memIncome6 = memIncome6 + shelf.__getitem__('month6')
cmonth6 = shelf.__getitem__('cmonth6')
cmonth5 = shelf.__getitem__('cmonth5')
cmonth4 = shelf.__getitem__('cmonth4')
cmonth3 = shelf.__getitem__('cmonth3')
cmonth2 = shelf.__getitem__('cmonth2')
cmonth1 = shelf.__getitem__('cmonth1')

memName = []
memName = memName + shelf.__getitem__('name')
print(memIncome2)
choice = int(input('Enter a member number:\n'))
members = memNum.index(choice)

income1 = None
income2 = None
income3 = None
income4 = None
income5 = None
income6 = None

try:
    income1 = memIncome1[members]
    income2 = memIncome2[members]
    income3 = memIncome3[members]
    income4 = memIncome4[members]
    income5 = memIncome5[members]
    income6 = memIncome6[members]
except IndexError:
    print('not all months have been set')
name = memName[members]
print(income1)
incomes = []
months = []
if income6 or income6 == 0:
    print(income6)
    incomes.append(income6)
    months.append(cmonth6)
if income5 or income5 == 0:
    print(income5)
    incomes.append(income5)
    months.append(cmonth5)
if income4 or income4 == 0:
    print(income4)
    incomes.append(income4)
    months.append(cmonth4)
if income3 or income3 == 0:
    print(income3)
    incomes.append(income3)
    months.append(cmonth3)
if income2 or income2 == 0:
    print(income2)
    incomes.append(income2)
    months.append(cmonth2)
incomes.append(income1)
months.append(cmonth1)
print(months)
print(incomes)
print(name)
draw_graph(incomes, months, name)