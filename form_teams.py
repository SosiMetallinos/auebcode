import argparse
import pandas as pd
import sys
from openpyxl import Workbook
from copy import deepcopy

#cd C:\Users\Sosipatros\Documents\GitHub\auebcode
#python form_teams.py students.xlsx

parser = argparse.ArgumentParser(description='Team Formation Program')
parser.add_argument('input', type=str, help='Input Excel file')
parser.add_argument('-o', '--output', type=str, help='Output Excel file')
args = parser.parse_args()  

def is_number(value):
    try:
        # Try converting to numeric
        numeric_value = pd.to_numeric(value)
        # Explicitly return False if the value is NaN
        if pd.isna(numeric_value):
            return False
        return True
    except ValueError:
        return False

class Student:
    def __init__(self, id, gender, score, friendid):
        self.id = id
        self.gender = gender
        self.score = score
        self.friendid = friendid
        self.category = 0
        self.team = 0

# Read input Excel file
input_data = pd.read_excel(args.input)
data_list = input_data.values.tolist()
sorted_data = sorted(data_list, key=lambda x: x[2])

# Calculate boundaries
n = len(sorted_data)
b1 = n // 4   # First boundary: 25th percentile
b2 = n // 2   # Second boundary: 50th percentile (median)
b3 = 3 * n // 4   # Third boundary: 75th percentile

count_males = 0
count_notmales = 0
all_students = []
for i in range(len(sorted_data)):
    if sorted_data[i][1] != 'male' and sorted_data[i][1] != 'female':
        sorted_data[i][1] = 'not-specified'
    
    if sorted_data[i][1] == 'male':
        count_males += 1
    else:
        count_notmales += 1
    
    if not(is_number(sorted_data[i][2])):
        sorted_data[i][2] = 0
    else:
        sorted_data[i][3] = sorted_data[i][3]
    
    if not(is_number(sorted_data[i][3])):
        sorted_data[i][3] = 0
    else:
        sorted_data[i][3] = int(sorted_data[i][3])
    
    student = Student(sorted_data[i][0], sorted_data[i][1], sorted_data[i][2], sorted_data[i][3])
    if i<b1:
        student.category = 1
    elif i<b2:
        student.category = 10 #assigns one of the 4 categories. Condition works because sorted_data is sorted based on score
    elif i<b3:
        student.category = 100
    else:
        student.category = 1000
    all_students.append(student)
    print(student.id, student.gender, student.score, student.friendid, student.category)

def assign_based_on_gender(students, m, ml, fm, ct): #Students DONT get to choose any member of their team, 2+ women per team of 4-5, balanced skills
    count_members = m               
    count_males = ml
    count_females = fm
    current_team = ct
    for i in range(len(students)):
        if students[i].team != 0: #Skip if team already assigned to this student
            continue
        if count_members >= 4: 
            current_team +=1
            count_members = 1
            if students[i].gender == 'male':
                count_males = 1
                count_females = 0
            else:
                count_females = 1
                count_males = 0
            students[i].team = current_team
        else:
            if students[i].gender == 'male' and count_males >=2:
                continue
            elif students[i].gender == 'female' and count_females >=2:
                continue
            else:
                students[i].team = current_team
                if students[i].gender == 'male':
                    count_males +=1
                else:
                    count_females +=1
                count_members +=1
            

            


    return 0

print(assign_based_on_gender(all_students))
print('Males:', count_males, 'Females:', count_notmales)

"""
-----------------------------------------------------------------------------------------------------------------------------
def exists_already(id, matrix):
    flag = False
    for i in range(len(matrix)):
        for j in matrix[i]:
            if (j == id):
                flag = True
                break
    return flag

row0 = [0, 0, ""]
pairs = [row0]
singles = [row0]
for i in range(len(data_list)):
    if (data_list[i][3] != None):
        if (not exists_already(data_list[i][0], pairs)):
            pairs.append([data_list[i][0], data_list[i][3], data_list[i][1]])
    else:
        if (not exists_already(data_list[i][0], singles)):
            singles.append([data_list[i][0], data_list[i][3], data_list[i][1]])

pairs.pop(0)
singles.pop(0)

if (len(pairs)<2):
    print("not enough pairs")
    sys.exit()

groups = [[pairs[0], pairs[1], 0, 0, 0, 0, 0]]

if (pairs[0][2]=="male"):
    groups[0][5] += 1
else:
    groups[0][6] += 1

for i in range(len(data_list)):
    gender = ""
    if (data_list[i][0]==pairs[0][1]):
        gender=data_list[i][1]

for i in range(1, len(pairs)):
    groups.append([0, 0, 0, 0, 0, 0, 0])
    groups[i][0] = pairs[i][0]
    groups[i][1] = pairs[i][1]
    if (pairs[i][2]=="male"):
        groups[i][5] += 1
    else:
        groups[i][6] += 1
    if (pairs[i][2]=="male"):
        groups[i][5] += 1
    else:
        groups[i][6] += 1

for i in range(len(groups)//2):
    if (groups[i][0] != pairs[i][0]):
        groups[i][2] = pairs[-1][0]
        groups[i][3] = pairs[-1][1]
        if (pairs[-1][0]=="male"):
            groups[i][5] += 1
        else:
            groups[i][6] += 1
        if (pairs[-1][1]=="male"):
            groups[i][5] += 1
        else:
            groups[i][6] += 1
        pairs.pop(-1)

for single in singles:
    for group in groups:
        if (group[5] + group[6] < 5):
            group[4] = single[0]
            if (single[0]=="male"):
                group[5] += 1
            else:
                group[6] += 1
            break
            
wb = Workbook()
ws = wb.active

for row in groups:
    ws.append([str(elem) for elem in row])

# Save the workbook
wb.save('teams.xlsx')
"""

print("Team formation completed successfully!")
