import argparse
import pandas as pd
import sys
from openpyxl import Workbook
from copy import deepcopy
import time

#python form_teams.py students.xlsx

print("Processing", end="") #Loading screen
for _ in range(4):
    sys.stdout.write('.')
    sys.stdout.flush()
    time.sleep(1)

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
    
def swap_gender(gender):
    if gender == 'male':
        return 'female'
    else:
        return 'male'

class Student:
    def __init__(self, id, gender, score, friendid):
        self.id = id
        self.gender = gender
        self.score = score
        self.friendid = friendid
        self.category = 0
        self.team = 0

class Team:

    def __init__(self, id, members):
        self.id = id
        # Ensure members is always a list
        if isinstance(members, list):
            self.members = members
        else:
            if members != 0:
                self.members = [members] 
            else:
                self.members = []
        self.males = 0
        self.females = 0
        self.scorebalance = 0
        for member in self.members:
            if member.gender == 'male':
                self.males +=1
            else:
                self.females +=1
            self.scorebalance += member.category
    
    def number_of(self, gender):
        if gender == 'male':
            return self.males
        else:
            return self.females

    def size(self):
        return len(self.members)
    
    def add_member(self, newmember):
        self.members.append(newmember)
        if newmember.gender == 'male':
            self.males += 1
        else:
            self.females += 1
        self.scorebalance += newmember.category

#Check for header
preview_data = pd.read_excel(args.input, nrows=5)
if all(isinstance(x, str) for x in preview_data.columns):
    # Read the entire file with the first row as the header
    input_data = pd.read_excel(args.input)
    data_list = input_data.values.tolist()
else:
    # Read the entire file without treating the first row as the header
    input_data = pd.read_excel(args.input, header=None)
    data_list = input_data.values.tolist()


sorted_by_id = sorted(data_list, key=lambda x: x[0])
for i in range(len(sorted_by_id) - 1):
    if not is_number(sorted_by_id[i][0]):
        print("ERROR: Missing students. Row:", i) #LOGIC ISSUE
        sys.exit(1)
    if sorted_by_id[i][0]==sorted_by_id[i+1][0]:
        j = i+1
        print("ERROR: Duplicate students found. Rows:", i, j) #LOGIC ISSUE
        sys.exit(1)


sorted_by_score = sorted(data_list, key=lambda x: x[2])

# Calculate boundaries
n = len(sorted_by_score)
b1 = n // 4   # First boundary: 25th percentile
b2 = n // 2   # Second boundary: 50th percentile (median)
b3 = 3 * n // 4   # Third boundary: 75th percentile

count_males = 0
count_notmales = 0
all_students = []
for i in range(len(sorted_by_score)):
    if sorted_by_score[i][1] != 'male' and sorted_by_score[i][1] != 'female': #fixes value issues with gender, primarly NaN=no value
        sorted_by_score[i][1] = 'not-specified'
    
    if sorted_by_score[i][1] == 'male':
        count_males += 1
    else:
        count_notmales += 1
    
    if not(is_number(sorted_by_score[i][2])): #fixes value issues with score, primarly NaN
        sorted_by_score[i][2] = 0
    else:
        sorted_by_score[i][3] = sorted_by_score[i][3]
    
    if not(is_number(sorted_by_score[i][3])): #fixes value issues with friend id, primarly NaN
        sorted_by_score[i][3] = 0
    else:
        sorted_by_score[i][3] = int(sorted_by_score[i][3])
    
    student = Student(sorted_by_score[i][0], sorted_by_score[i][1], sorted_by_score[i][2], sorted_by_score[i][3])
    if i<b1:
        student.category = 1 #assigns one of the 4 categories. Condition works because sorted_by_score is sorted based on score
    elif i<b2:
        student.category = 10 
    elif i<b3:
        student.category = 100
    else:
        student.category = 1000
    all_students.append(student)


#INITIALIZE VARIABLES
remaining = {}
remaining['male'] = count_males  #at some point we will need to know if we have run out of one gender
remaining['female'] = count_notmales
new_team_id = 1
all_students[0].team = new_team_id
initial_members = [ all_students[0] ]
team1 = Team(new_team_id, initial_members)
remaining[all_students[0].gender] -= 1
#print('id:', team1.id, team1.members[0].id, 'size', team1.size, 'males', team1.males, 'females', team1.females, 'scb', team1.scorebalance)
all_teams = [team1]
for index, student in enumerate(all_students):
    if student.team != 0:
        remaining[student.gender] -= 1
        continue
    entered = False
    for team in all_teams:
        if team.number_of(student.gender)<2 and team.size()<4:
            entered = True
            remaining[student.gender] -= 1
            student.team = team.id
            team.add_member(student)
            break

    #For very bad data e.g. too many consecutive men 
    if not entered:
        for i in range(index+1, len(all_students)):
            next_gender_ctgr = 10000
            if all_students[i].gender != student.gender and all_students[i].team == 0:
                if i < index+8:
                    break
                next_gender_ctgr = all_students[i].category
            else:
                continue
            if next_gender_ctgr <= student.category*10: #It is implied that i>=index+8 as well
                for each_team in all_teams:
                    if each_team.size() <4 and each_team.males*each_team.females==0: #not a full team and the other gender is zero
                        entered = True
                        remaining[student.gender] -=1
                        student.team = each_team.id
                        each_team.add_member(student)
                        break
            if entered:
                break
    
    if not entered: #all attempts to put him/her in an existing team failed. Now let's make a new one
        new_team_id += 1
        student.team = new_team_id
        initial_members = [student]
        new_team = Team(new_team_id, initial_members)
        all_teams.append(new_team)


iteratable_teams = []
for team in all_teams:
    memberlist = [team.id]  # Start with the team ID
    for member in team.members:
        memberlist.append(member.id)
    memberlist.append(team.scorebalance)
    memberlist.append(team.males)
    memberlist.append(team.females)
    iteratable_teams.append(memberlist)
wb = Workbook()
ws = wb.active

headers = ["Team ID", "Member1", "Member2", "Member3", "Member4", "Balance", "Males", "Females"]
ws.append(headers)
for team in iteratable_teams:
    print(team)
    ws.append([str(elem) for elem in team])

# Save the workbook
wb.save('teams.xlsx')

print("\n\nDone!")
print("You will find a teams.xlsx file in the same folder as the script. Previous teams.xlsx files will be overwritten.")


"""
PREVIOUS ATEMPT

def assign_based_on_gender(students, m, ml, fm, ct, teamed, teams): #Students DONT get to choose any member of their team, 2+ women per team of 4-5, balanced skills
    count_members = m               
    count_males = ml
    count_females = fm
    current_team = ct
    for i in range(len(students)):
        if students[i].team != 0: #Skip if student already assigned a team
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
        teamed +=1 #INCORRECT RESULT, WHY?
        if current_team in teams:
            teams[current_team].append(students[i].id)
        else:
            teams[current_team] = [students[i].id]
            
    return students, count_members, count_males, count_females, current_team, teamed, teams

repeat = 0
stud, cmm, cml, cfm, ct, teamed, teams = deepcopy(all_students), 0, 0, 0, 1, 0, {}
while teamed <= len(all_students) and repeat<=len(all_students)-teamed:
    repeat +=1
    stud, cmm, cml, cfm, ct, teamed, teams = assign_based_on_gender(deepcopy(stud), cmm, cml, cfm, ct, teamed, deepcopy(teams))

print('Repetitions:', repeat)
print(cmm, cml, cfm, ct, teamed, len(teams))
for key in teams:
    print(key, teams[key])
"""