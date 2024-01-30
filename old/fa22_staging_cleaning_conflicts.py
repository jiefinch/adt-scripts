#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Dec  5 12:23:09 2022

@author: cindywang

hanjie@adt-scripts.iam.gserviceaccount.com

"""

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import copy


# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
# THINGS TO EDIT PER SEMESTER ------------------------------------------------------------
spreadsheet = client.open("[Spring 2023] Prod Week Conflicts (Responses)")

responses_sheet = spreadsheet.get_worksheet(1)

cleaning_conflicts_matrix_sheet = spreadsheet.get_worksheet(9)
cleaning_excuses_matrix_sheet = spreadsheet.get_worksheet(10)
staging_conflicts_matrix_sheet = spreadsheet.get_worksheet(11)
staging_excuses_matrix_sheet = spreadsheet.get_worksheet(12)

first_response_row = 2
last_response_row = 213
dances_str = "Spring Fragrance, Lotus in June, Tang Impressions, Like the Sunshine, Bathing in Heaven, Unfettered, Water, D.D.D, Flower, Cereal Dreams, DM, Full Moon, Queendom, Chain, Halazia, Now & Forever, Secret Story of the Swan, 28 Reasons, Senior Interlude"
# ------------------- DONT TOUCH BELOW ---------------------------------------------

responses_list_of_lists = responses_sheet.get_all_values()
cleaning_conflicts_matrix_sheet.batch_clear(['A10:Z500'])
cleaning_excuses_matrix_sheet.batch_clear(['A10:Z500'])
staging_conflicts_matrix_sheet.batch_clear(['A10:Z500'])
staging_excuses_matrix_sheet.batch_clear(['A10:Z500'])

dances = dances_str.split(', ')
print(dances)

times_str = "8:00 - 8:30 AM, 8:30 - 9:00 AM, 9:00 - 9:30 AM, 9:30 - 10:00 AM, 10:00 - 10:30 AM, 10:30 - 11:00 AM, 11:00 - 11:30 AM, 11:30 - 12:00 PM, 12:00 - 12:30 PM, 12:30 - 1:00 PM, 1:00 - 1:30 PM, 1:30 - 2:00 PM, 2:00 - 2:30 PM, 2:30 - 3:00 PM, 3:00 - 3:30 PM, 3:30 - 4:00 PM, 4:00 - 4:30 PM, 4:30 - 5:00 PM, 5:00 - 5:30 PM, 5:30 - 6:00 PM, 6:00 - 6:30 PM, 6:30 - 7:00 PM, 7:00 - 7:30 PM, 7:30 - 8:00 PM, 8:00 - 8:30 PM, 8:30 - 9:00 PM, 9:00 - 9:30 PM, 9:30 - 10:00 PM, 10:00 - 10:30 PM, 10:30 - 11:00 PM, 11:00 - 11:30 PM, 11:30 - 12:00 AM"
times = times_str.split(', ')

print("time strings", times)

#sorry brain dead this is spaghetti code so sorry
cleaning_conflict_matrix = [[t] for t in times]
none = [None] * len(dances)
for row in cleaning_conflict_matrix:
    row.extend([None]*len(dances))
header = ['']
header.extend(dances)
header = [header]
header.extend(cleaning_conflict_matrix)
cleaning_conflict_matrix = header
print(cleaning_conflict_matrix)

dance_ind = {}
for dance in dances:
    dance_ind[dance] = cleaning_conflict_matrix[0].index(dance)

time_ind = {}
for i in range(1, len(cleaning_conflict_matrix)):
    time_ind[cleaning_conflict_matrix[i][0]] = i
    
cleaning_excuse_matrix = copy.deepcopy(cleaning_conflict_matrix)
staging_conflict_matrix = copy.deepcopy(cleaning_conflict_matrix)
staging_excuse_matrix = copy.deepcopy(cleaning_conflict_matrix)
    

for row_ind in range(first_response_row,last_response_row):
    row = responses_list_of_lists[row_ind]
    name = row[1]
    relevant_dances = row[3].split(", ")
    cleaning_conflict_times = row[4].split(", ")
    cleaning_excuse = row[5]
    staging_conflict_times = row[6].split(", ")
    staging_excuse = row[7]
    
    for d in relevant_dances:
        for ct in cleaning_conflict_times:
            if ct == 'N/A':
                continue
            if not cleaning_conflict_matrix[time_ind[ct]][dance_ind[d]]:
                cleaning_conflict_matrix[time_ind[ct]][dance_ind[d]] = 0
                cleaning_excuse_matrix[time_ind[ct]][dance_ind[d]] = {
                    'names': [], 
                    'excuses': [],
                    }
            cleaning_conflict_matrix[time_ind[ct]][dance_ind[d]] += 1
            cleaning_excuse_matrix[time_ind[ct]][dance_ind[d]]['names'].append(name)
            cleaning_excuse_matrix[time_ind[ct]][dance_ind[d]]['excuses'].append(cleaning_excuse)
            
        for st in staging_conflict_times:
            if st == 'N/A':
                continue
            #print("LOOKING AT ", st, "IN", time_ind)
            if not staging_conflict_matrix[time_ind[st]][dance_ind[d]]:
                staging_conflict_matrix[time_ind[st]][dance_ind[d]] = 0
                staging_excuse_matrix[time_ind[st]][dance_ind[d]] = {
                    'names': [], 
                    'excuses': [],
                    }
            staging_conflict_matrix[time_ind[st]][dance_ind[d]] += 1
            staging_excuse_matrix[time_ind[st]][dance_ind[d]]['names'].append(name)
            staging_excuse_matrix[time_ind[st]][dance_ind[d]]['excuses'].append(staging_excuse)
            
for r, row in enumerate(cleaning_excuse_matrix):
    for c, entry in enumerate(row):
        if not entry:
            cleaning_excuse_matrix[r][c] = ''
        else:
            cleaning_excuse_matrix[r][c] = str(entry)
            
for r, row in enumerate(staging_excuse_matrix):
    for c, entry in enumerate(row):
        if not entry:
            staging_excuse_matrix[r][c] = ''
        else:
            staging_excuse_matrix[r][c] = str(entry)
    
cleaning_conflicts_matrix_sheet.update('A1:Z35', cleaning_conflict_matrix)
cleaning_excuses_matrix_sheet.update('A1:Z35', cleaning_excuse_matrix)
staging_conflicts_matrix_sheet.update('A1:Z35', staging_conflict_matrix)
staging_excuses_matrix_sheet.update('A1:Z35', staging_excuse_matrix)

    
    