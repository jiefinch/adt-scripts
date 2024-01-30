#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Oct 31 12:49:25 2022

@author: cindywang
"""

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Apr 26 11:50:58 2022

@author: cindywang
"""
import sys, subprocess
subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'gspread'])
subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'oauth2client'])



import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
from time import strftime, localtime
import copy


# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
spreadsheet = client.open("[Fall 2023] Master Roster")

master_roster_sheet = spreadsheet.get_worksheet(0)
master_roster_list_of_lists = master_roster_sheet.get_all_values()

conflict_matrix_worksheet = spreadsheet.get_worksheet(1)
# print(master_roster_list_of_lists)

dances_start_ind = 7
dances_end_ind = len(master_roster_list_of_lists[0])
dances = master_roster_list_of_lists[0][dances_start_ind:dances_end_ind]
print(dances)



#sorry brain dead this is spaghetti code so sorry
conflict_matrix = [[d] for d in dances]
print(conflict_matrix)
print(len(dances))
none = [None] * len(dances)
print(none)
for row in conflict_matrix:
    row.extend([None]*len(dances))
header = ['']
header.extend(dances)
header = [header]
header.extend(conflict_matrix)
conflict_matrix = header
print(conflict_matrix)

dance_ind = {}
for dance in dances:
    dance_ind[dance] = conflict_matrix[0].index(dance)

first_dancer_row = 2
last_dancer_row = 247 # TODO UPDATE THIS EVERY TIME
for row in range(first_dancer_row,last_dancer_row-1):
    dancer = ''
    if master_roster_list_of_lists[row][1]:
        dancer = master_roster_list_of_lists[row][0] + ' ' + master_roster_list_of_lists[row][1] + ' ' + master_roster_list_of_lists[row][2]
    else:
        dancer = master_roster_list_of_lists[row][0] + ' ' + master_roster_list_of_lists[row][2]
        
    dances = []
        
    for col in range(dances_start_ind, dances_end_ind):
        if master_roster_list_of_lists[row][col].strip().lower() == 'x':
            dances.append(master_roster_list_of_lists[0][col])
    
    if len(dances) < 2:
        continue
    for i in range(len(dances)):
        for j in range(i+1, len(dances)):
            dance1_ind, dance2_ind = dance_ind[dances[i]], dance_ind[dances[j]]
            if dance1_ind < dance2_ind:
                dance1_ind, dance2_ind = dance2_ind, dance1_ind
                
            if not conflict_matrix[dance1_ind][dance2_ind]:
                conflict_matrix[dance1_ind][dance2_ind] = {
                    'dancers': [],
                    'num_conflicts': 0
                    }
            conflict_matrix[dance1_ind][dance2_ind]['dancers'].append(dancer)
            conflict_matrix[dance1_ind][dance2_ind]['num_conflicts']+=1
            
print(conflict_matrix)

conflict_matrix_copy = copy.deepcopy(conflict_matrix)

for r, row in enumerate(conflict_matrix):
    for c, entry in enumerate(row):
        if not entry:
            conflict_matrix[r][c] = ''
        else:
            if r > 0 and c > 0:
                conflict_matrix[r][c] = int(entry['num_conflicts'])
        

conflict_matrix_worksheet.update('A1:Z30',conflict_matrix)


for r, row in enumerate(conflict_matrix_copy):
    for c, entry in enumerate(row):
        if not entry:
            conflict_matrix_copy[r][c] = ''
        else:
            conflict_matrix_copy[r][c] = str(entry)
            
conflict_matrix_worksheet.update('A24:Z60',conflict_matrix_copy)
            


# Extract and print all of the values
# list_of_hashes = responses_sheet.get_all_records()
# print(list_of_hashes)

# dance_col_key = {
#     "Jopping": 2,
#     "Spring Wind": 3,
#     "Mountain Spirits": 4,
#     "Bad Love": 5,
#     "SILK": 6,
#     "pretty girls are my besties": 7,
#     "Dreaming Night": 8,
#     "Crimson Lips": 9,
#     "We Must Love": 10,
#     "Step on Me": 11,
#     "Silent Sky": 12,
#     "Hello My Future": 13,
#     "Daybreak": 14,
#     "Hula Hoop": 15,
#     "Galloping": 16,
#     "Play Me Kill Me": 17,
#     "Nature's Blade": 18,
#     "Reveal": 19,
#     "Senior Interlude": 20
#     }

# def strip_email(inp_email):
#     ind = inp_email.find('@')
#     if ind > 0:
#         return inp_email[0:ind].lower().strip()
#     return inp_email.lower().strip()

# # need to preprocess to delete invalid early attestations
# def get_attestestations_per_day(date, sheet):
#     responses_sheet = spreadsheet.get_worksheet(0)
#     master_roster = spreadsheet.get_worksheet(1)
#     attestation_results = spreadsheet.get_worksheet(sheet)
    

#     attestation_results.batch_clear(['A10:E500'])
#     emails = master_roster.col_values(4)[1:]

#     kerbs = set()

#     for e in emails:
#         ind = e.find('@')
#         if ind > 0:
#             kerbs.add(e[0:ind].lower().strip())
#         else:
#             print(e + ' does not look like an email')
            
#     attestations_list_of_lists = responses_sheet.get_all_values()
    
#     attested = set()
    
#     not_attending = []
    
#     for row_ind in range(1, len(attestations_list_of_lists)):
#         if date in attestations_list_of_lists[row_ind][4]:
#             if attestations_list_of_lists[row_ind][5] != 'Yes':
#                 not_attending.append({
#                     'name': attestations_list_of_lists[row_ind][1],
#                     'kerb': strip_email(attestations_list_of_lists[row_ind][2]),
#                     'dances': attestations_list_of_lists[row_ind][3].split(", ")
#                     })
#             attested.add(strip_email(attestations_list_of_lists[row_ind][2]))
    
#     dances = list(dance_col_key.keys())
    
#     result_list_of_lists = [dances]
#     absent_per_dance = []
        
#     for c in range(0, len(dances)):
#         absent_this_dance = ''
#         for a in not_attending:
#             if dances[c] in a['dances']:
#                 absent_this_dance += a['name'] + ', '
#         absent_per_dance.append(absent_this_dance)
        
#     result_list_of_lists.append(absent_per_dance)
    
#     did_not_fill = list(kerbs - attested)
#     did_not_fill_list_of_lists = [[a] for a in did_not_fill]
#     did_not_fill_list_of_lists.insert(0, ["Did not submit:"])
    
#     filled_but_shouldnt = list(attested - kerbs)
#     filled_but_shouldnt_list_of_lists = [[a] for a in filled_but_shouldnt]
#     filled_but_shouldnt_list_of_lists.insert(0, ["Not on master roster:"])
        
#     attestation_results.update('B1:Z5',result_list_of_lists)
    
#     master_roster_all = master_roster.get_all_values()
#     kerb_to_name = {}
    
#     for row in range(1, len(master_roster_all)):
#         name = master_roster_all[row][0]
#         if master_roster_all[row][1]:
#             name += " (" + master_roster_all[row][1] + ")"
            
#         name += " " + master_roster_all[row][2]
#         kerb_to_name[strip_email(master_roster_all[row][3])] = name
        
#     for i in range(1, len(did_not_fill_list_of_lists)):
#         did_not_fill_list_of_lists[i].append(kerb_to_name[did_not_fill_list_of_lists[i][0]])
#         did_not_fill_list_of_lists[i].append(did_not_fill_list_of_lists[i][0] + "@mit.edu")

#     attestation_results.update('A10:C500',did_not_fill_list_of_lists)
#     attestation_results.update('E10:E500',filled_but_shouldnt_list_of_lists)
    
#     public_sheet = spreadsheet2.get_worksheet(sheet-2)
#     public_sheet.batch_clear(['A10:E500'])
#     public_sheet.update('B1:Z5',result_list_of_lists)
#     public_sheet.update('A10:C500',did_not_fill_list_of_lists)
#     public_sheet.update('E10:E500',filled_but_shouldnt_list_of_lists)

# # get_attestestations_per_day("May 7", 2)
# # get_attestestations_per_day("May 8", 3)
# while True:
#     # get_attestestations_per_day("May 9", 4)
#     # get_attestestations_per_day("May 10", 5)
#     get_attestestations_per_day("May 11", 6)
#     get_attestestations_per_day("May 12", 7)
#     print(strftime("%Y-%m-%d %H:%M:%S", localtime()))
#     time.sleep(600)
            
        
# # keep set of kerbs that submitted for a given day
# # make list of kerbs that answered no
# # compare submitted to set of all kerbs, find difference
# # put list of kerbs that did not submit or answered no onto a sheet

# #clear the conflict matrix every time you run the script
