import gspread
import pprint
import os
from oauth2client.service_account import ServiceAccountCredentials
from difflib import SequenceMatcher
import time

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

CREDS_FILE = r'C:\Users\chris\Videos\Local Top Spec\Software\Archive Checker\creds.json'
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, scope)

client = gspread.authorize(creds)

sheet = client.open('Top Spec Schedule').get_worksheet(0)
maxRowCount = sheet.row_count
#print(maxRowCount)

#cell = str(sheet.get('C95'))
#print(similar("Focusrite Scarlett 2i2 Unboxing", "Macbook Pro 16 Unboxing"))

currentVidNum = 19

def update_archive():
    for row in range(9, maxRowCount):
        currentVidStatus = sheet.get('E' + str(row))
        print(currentVidStatus[0][0])
        if(currentVidStatus[0][0] == 'D'):
            currentVidNum = sheet.get('I' + str(row))
            print("Checking for video number " + str(currentVidNum))
            notFound = True
            isArchived = ""
            for currentDir in dir_contents:
                print("Checking " + currentDir + "For " + str(currentVidNum[0][0]))
                currentDirNum = currentDir[0:3]
                if currentDirNum.__contains__(currentVidNum[0][0]):
                    isArchived = "=CHAR(10004)"
                    break
                else:
                    isArchived = "=CHAR(10006)"
            sheet.update_cell(row, 8, isArchived)
            time.sleep(3.0)
        elif currentVidStatus[0][0] == '?':
            sheet.update_cell(row, 8, "-")

ARCHIVE_PATH = r'T:\Archive'
if(os.path.isdir(ARCHIVE_PATH) == True):
    dir_contents = os.listdir(ARCHIVE_PATH)
    print("Mounted to Archive")
    print(*dir_contents, sep = "\n")
    update_archive()
else:
    print("Archive Not Connected")