
from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

from datetime import datetime
from dateutil import relativedelta

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

#grabs values in a google spreadsheet
def get_sheet(sheetId, range):
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheetId = sheetId
    rangeName = range
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    return result.get('values', [])

#sets values in a google spreadsheet
def set_sheet(sheetId, range,body):
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheetId = sheetId
    rangeName = range
    result = service.spreadsheets().values().update(
        spreadsheetId=spreadsheetId, range=rangeName, valueInputOption='USER_ENTERED',body=body).execute()
    return result

#appends values in a google spreadsheet
def update_sheet(sheetId, range,body): 
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheetId = sheetId
    rangeName = range
    result = service.spreadsheets().values().append(
        spreadsheetId=spreadsheetId, range=rangeName, valueInputOption='RAW',
        insertDataOption='INSERT_ROWS', body=body).execute()
    return result

#does a percent match of the mentor answers compared to the mentee answers
def percent_match(menteeArray, mentorArray):
    menteeSplit = menteeArray.split(',') 
    matchCounter = 0.0
    menteeTotal = 0.0
    for each in menteeSplit:
        menteeTotal = menteeTotal + 1
        if each in mentorArray:
            matchCounter = matchCounter + 1
    return matchCounter / menteeTotal

#identifies if at least one of the mentees answers is in the mentors answers
def one_match(menteeArray, mentorArray):
    menteeSplit = menteeArray.split(',') 
    match = 0
    for each in menteeSplit:
        if each in mentorArray:
            match = 1
            break
    return match

#Excutes the main mentorship matching processs
def main():

    #get the last time the process was ran
    lastRuntimeValues = get_sheet('15pirsOaBtQ5gJgw_ZFGuBx5kGQmuWZcZb-AN8feCgfQ','Sheet1!A:A')
    if not lastRuntimeValues:
        print('Last runtime was not returned')

    #setting last runtime
    lastRuntimeIndex = len(lastRuntimeValues)
    lastRuntimeList = lastRuntimeValues[len(lastRuntimeValues) - 1]
    lastRuntime = datetime.strptime(lastRuntimeList[0], "%m/%d/%Y %H:%M:%S")

    #append the current time in the last run time sheet
    body = {'values':[[datetime.now().strftime("%m/%d/%Y %H:%M:%S")]]}
    #result = update_sheet('15pirsOaBtQ5gJgw_ZFGuBx5kGQmuWZcZb-AN8feCgfQ','Sheet1!A:A',body)
    #if not result:
       # print('Runtime was not appended')

    #get values from mentorship from
    mentorshipFormValues = get_sheet('1HGXZS72XRe5Q5N-qrZHVUU7krBjGplFqA3XYUVmU_LQ','Form Responses 1!A2:AH')
    if not mentorshipFormValues:
        print('No mentorship form data returned.')
    else:
        mentorDict = {}
        menteeDict = {}

        for row in mentorshipFormValues:
            #add new users after the last runtime stamp and
            #determine if the user wants to be a mentor/mentee/both and add them to the associated array

            if datetime.strptime(row[0], "%m/%d/%Y %H:%M:%S") > lastRuntime:
                if row[14] == 'Mentor':
                    mentorDict[row[3]] = [row[0], row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10],row[11],row[12],row[13],row[26],row[27],row[28],row[29],row[30],row[31],row[32],row[33]]

                elif row[14] == 'Mentee':
                    menteeDict[row[3]] = [row[0], row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10],row[11],row[20],row[21],row[22],row[23],row[24],row[25],row[31],row[32],row[33]]
                elif row[14] == 'Both':
                    menteeDict[row[3]] = [row[0], row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10],row[11],row[20],row[21],row[22],row[23],row[24],row[25],row[31],row[32],row[33]]
                    mentorDict[row[3]] = [row[0], row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10],row[11],row[12],row[13],row[15],row[16],row[17],row[18],row[19],row[31],row[32],row[33]]

                else:
                    print('Not denoted as mentor/mentee/both. Email: '+row[3]+ ' value: '+row[14])
    
    #get mentee waitlist to see if any new mentees are already on list
    menteeWaitlistValues = get_sheet('1vxqr3gdO-etFuOI7FV1HlVPr1q4r9fS0Mra3ExX_L6w','Sheet1!A2:V')
    if not menteeWaitlistValues:
        print('No mentee waitlist data returned.')

    for idx, row in enumerate(menteeWaitlistValues):
        #checking if the email is in the new mentee dictionary and the mentee has not been matched
        if row[3] in menteeDict and len(row) == 21:

            #updating the timestamp to be the original sign up date
            newMenteeValues = menteeDict[row[3]]
            newMenteeValues[0] = row[0]

            #update mentee in waitlist
            body = {'values': [newMenteeValues]}
            result = set_sheet('1vxqr3gdO-etFuOI7FV1HlVPr1q4r9fS0Mra3ExX_L6w','Sheet1!A'+str(idx+2)+':U',body)
            if not result:
                print('ERROR update mentee waitlist. Email: '+row[3])
           
            #remove mentee from menteeDict
            else:
                del menteeDict[row[3]]

    #adding the remaining new mentee values to an array
    newMenteeArray = []
    for key in menteeDict:
        newMenteeArray.append(menteeDict[key])
    
    #append the new mentees to mentee waitlist
    body = {'values':newMenteeArray}
    result = update_sheet('1vxqr3gdO-etFuOI7FV1HlVPr1q4r9fS0Mra3ExX_L6w','Sheet1!A2:U',body)
    if not result:
       print('Mentees was not appended to waitlist')

    #get all mentees from waitlist
    menteeWaitlistValues = get_sheet('1vxqr3gdO-etFuOI7FV1HlVPr1q4r9fS0Mra3ExX_L6w','Sheet1!A2:V')
    if not menteeWaitlistValues:
        print('No mentee waitlist data returned.')

    #removing any matched mentees
    for row in menteeWaitlistValues:
        if len(row) == 22:
            menteeWaitlistValues.remove(row)

    #MAGIC TIME! Loop through each mentor and apply match algorithm 
    matchArray = []
    for key in mentorDict:
        mentorArray = mentorDict[key]
        for menteeArray in menteeWaitlistValues:
            matchScore = 0

            #mentor must have more years than mentee
            if float(mentorArray[8]) < float(menteeArray[8]):
                #mentee has more years this is not a match
                continue 

            #must match on gender perference
            if menteeArray[12] != 'No preference, first mentor available':
                if mentorArray[12] != menteeArray[12]:
                    #gender preferences do not match 
                    continue

            #making sure position of mentee and mentor line up
            positionScore = 0
            
            if mentorArray[17] == menteeArray[17]:
                positionScore = 1
            elif mentorArray[17] == 'Entry':
                continue #entry mentors matched with entry mentees
            elif mentorArray[17] == 'Mid Level' and menteeArray[17] != 'Entry':
                continue #mid level mentors matched with mid level or entry mentees
            elif mentorArray[17] == 'Executive':
                positionScore = .75
            else:
                positionScore = .5
            #print(menteeArray[17]+' '+mentorArray[17])

            #time on waitlist
            menteeSignup = datetime.strptime(menteeArray[0], "%m/%d/%Y %H:%M:%S")
            todayDate = datetime.now()
            dateDiff = relativedelta.relativedelta(todayDate, menteeSignup)
            
            waitlistTime = dateDiff.months

            #personality
            personalityScore = float(percent_match(menteeArray[9],mentorArray[9]))

            #activites
            activityScore = percent_match(menteeArray[10],mentorArray[10])

            #values
            valueScore = percent_match(menteeArray[11],mentorArray[11])

            #desired support
            supportScore = one_match(menteeArray[14],mentorArray[14])

            #industry
            indsutryScore = one_match(menteeArray[15],mentorArray[15])

            #role
            if 'Unknown at this time' in menteeArray[16]:
                roleScore = 1
            else:
                roleScore = one_match(menteeArray[16],mentorArray[16])


            #communication
            communicationScore = one_match(menteeArray[18],mentorArray[19])

            #location
            locationScore = one_match(menteeArray[19],mentorArray[20])
            
            #time
            timeScore = one_match(menteeArray[20],mentorArray[21]) 

            matchScore = (waitlistTime*.05) + (communicationScore*.05) + (locationScore*.05) + (timeScore*.05) + (personalityScore*.05) + (activityScore*.05) + (valueScore*.05) + (supportScore*.25)+ (indsutryScore*.15) + (roleScore*.15) + (positionScore*.1)

            matchArray.append([str(matchScore),mentorArray[0],mentorArray[3],mentorArray[1],mentorArray[2],mentorArray[5],mentorArray[18],menteeArray[0],menteeArray[3],menteeArray[1],menteeArray[2],menteeArray[5]])

    #append matches to potential match sheet
    body = {'values':matchArray}
    result = update_sheet('1QVPlpdxvGyQxB8jf6zD_Ci9xro-aXxVbyiwkRb6Oito','Sheet1!A2:L',body)
    if not result:
       print('No potential matches appended')
    else:
        print('Potentail matches have been appended to file')

if __name__ == '__main__':
    main()
