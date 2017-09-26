# uiuan 2017-09-27
# v0.1 beta
# Listening part only
# Display points only
import pandas as pd
from pandas import ExcelFile
from pandas import ExcelWriter
import numpy as np

df = pd.read_excel('Data/Listening.xlsx', sheetname='Sheet1')

# print("Column headings:")
# print(df.columns)

# print("Values")
# for i in df.index:
#     print(df['Source'][i],df['No'][i])
# Welcome
print ('Now support Listening only')

# Load data
a_no = df['No']
a_c1 = df['Con1']
a_c2 = df['Con2']
a_l1 = df['Lec1']
a_l2 = df['Lec2']
a_l3 = df['Lec3']
a_l4 = df['Lec4']
no_max = np.max(a_no)
no_min = np.min(a_no)

print ('Load data from TPO', no_min, 'to', no_max, 'complete')

# Define the function to get Listening mark
def get_lmark( line, qt = 'all', file = 1 ):
    global a_c1,a_c2,a_l1,a_l2,a_l3,a_l4
    if file == 1:
        if qt == 'Con1':
            return int(a_c1[line])
        if qt == 'Con2':
            return int(a_c2[line])
        if qt == 'Lec1':
            return int(a_l1[line])
        if qt == 'Lec2':
            return int(a_l2[line])
        if qt == 'Lec3':
            return int(a_l3[line])
        if qt == 'Lec4':
            return int(a_l4[line])
        if qt == 'all':
            return (int(a_c1[line])+int(a_c2[line])+int(a_l1[line])+int(a_l2[line])+int(a_l3[line])+int(a_l4[line]))
        else:
            return 0
    if file == 0:
        return 0    

# Question info
qt_no = int(input('Type TPO number '))
if qt_no > no_max or qt_no < no_min:
    print ('No such TPO')
else:
    n = np.where(a_no==qt_no)[0]
    print ('The total score of listening part in TPO', qt_no ,'is',get_lmark(n))
    for i in range(1,3):
        print ('Conversation', i, 'has', get_lmark(n, 'Con'+str(i)), 'points')
    for i in range(1,5):
        print ('Lecture', i, 'has', get_lmark(n, 'Lec'+str(i)), 'points')