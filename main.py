from openpyxl import Workbook
from openpyxl import load_workbook
import warnings
from dataclasses import dataclass


@dataclass
class CellData:
    value: str
    color: str


'''
Excel has a feature called Data Validation where you can 
pick from a list of rules to limit the type of data that 
can be entered in a cell. This is sometimes used to create 
dropdown lists in Excel. This warning is telling you that 
this feature is not supported by openpyxl, and those rules 
will not be enforced.
'''
# Ignore warning
warnings.simplefilter(action='ignore', category=UserWarning)

wb1 = load_workbook('Alpha Test Report 2024-Henry.xlsm')
ws1 = wb1['Daily Tracker-Week-1']

# All the data that we need from the WS
data_range = ws1['A15':'AC59']

# Same data, just in a list
data_range_list = []

for lock in data_range:
    lock_list = []
    for cell in lock:
        print(f'Color: {cell.fill.fgColor.rgb}, type: {type(cell.fill.fgColor.rgb)}')
        lock_list.append(CellData(str(cell.value), cell.fill.fgColor.rgb))
        # print(cell.value, end=" ")
    # print()
    data_range_list.append(lock_list)


# Removing all locks that have a '.' in the name
# Adding the OpenLock values from current week
lock_dict = dict()
for row in data_range_list:
    if "." in row[0].value:
        continue
    else:
        # Cells 15-21 are labeled OpenLock
        lock_dict[row[0].value] = row[15:21]


# # Removing None values from the lists
# for lock, value in lock_dict.items():
#     l = value
#     for row in l:
#         for
#     lock_dict[lock] = l

'''
>>> print(cell11.fill.fgColor.rgb)
FFFFFF00
>>> type(cell11.fill.fgColor.rgb)
<class 'str'>
'''


# Print Dictionary
for lock, value in lock_dict.items():
    print(f'{lock} ', end='')
    for i in value:
        if i.value != None:
            print(f'{i.value}, {i.color} ', end='')
    print()

wb1.close()