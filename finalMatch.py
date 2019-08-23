
from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

from datetime import datetime
from dateutil import relativedelta
import hashlib
import requests
import json

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

#hashes email for MD5
def MD5(email):
    m = hashlib.md5()
    m.update(email)
    return m.hexdigest()

#sends emails to mailchimp
def send_email(email, payload, requestType):
    #setting variables for mailchimp
    listId = 'f09b66a085'
    userHash = MD5(email)
  
    #setting headers for API call
    headers = {
    'Authorization' : 'ENTER API KEY HERE',
    'content-type': "application/json"
    }

    url = 'https://us5.api.mailchimp.com/3.0/lists/'+listId+'/members/'+userHash
    if requestType == 'POST':
        url = 'https://us5.api.mailchimp.com/3.0/lists/'+listId+'/members'

    return requests.request(requestType, url, data=payload, headers=headers)

#Adds final matches to master match file and sends emails
def main():
    sendMentorEmailBool = 'TRUE'
    sendMenteeEmailBool = 'TRUE'

    today = datetime.now().strftime("%m/%d/%Y %H:%M:%S")

    finalMatchArray = []
    menteeEmailArray = []
    masterMentorDict = {}
    #get the potential matches 
    potentialMatchValues = get_sheet('1QVPlpdxvGyQxB8jf6zD_Ci9xro-aXxVbyiwkRb6Oito','Sheet1!A2:N')
    if not potentialMatchValues:
        print('Potential Matches not returned')
    else:
        #filter through the potential matches that final match is set to TRUE but email has not been sent
        for idx, row in enumerate(potentialMatchValues):
            if len(row) == 13 and row[12] == 'TRUE':

                #SENDING emails for mentee
                #Make payload for mentee
                menteeData = "{\n \"email_address\": \""+row[8]+"\", \"status_if_new\": \"subscribed\", \"merge_fields\": {\"FNAME\": \"" + row[9] + "\", \"MMERGE6\": \"" + row[3]+" "+row[4] + "\", \"MMERGE7\": \"" + row[2] + "\", \"MMERGE8\": \"" + row[5] + "\", \"MMERGE9\": \"Mentee\",\"MMERGE10\": \"FALSE\"}\n}"
     
                if sendMenteeEmailBool == 'TRUE':
                    sendMenteeEmail = send_email(row[8],menteeData,'PUT')

                    if sendMenteeEmail.status_code == 404:
                        #Make payload for mentee
                         menteeData = "{\n \"email_address\": \""+row[8]+"\", \"status\": \"subscribed\", \"merge_fields\": {\"FNAME\": \"" + row[9] + "\", \"MMERGE6\": \"" + row[3]+" "+row[4] + "\", \"MMERGE7\": \"" + row[2] + "\", \"MMERGE8\": \"" + row[5] + "\", \"MMERGE9\": \"Mentee\",\"MMERGE10\": \"FALSE\",\"MMERGE11\": \"TRUE\"}\n}"
                         sendMenteeEmail.status_code = send_email(row[8],menteeData,'POST')

                    #update menteeemail status in potential matches spreadsheet
                    body = {'values': [[sendMenteeEmail.status_code]]}
                    result = set_sheet('1QVPlpdxvGyQxB8jf6zD_Ci9xro-aXxVbyiwkRb6Oito','Sheet1!O'+str(idx+2),body)
                    if not result:
                        print('ERROR update potential match emails. Row Number: '+str(idx+2))
                    else:
                        print(menteeData)

                #Adding mentors to mentor master dictionary
                #Make payload for mentor
                if row[2] in masterMentorDict.keys():
                    if masterMentorDict[row[2]]['count']=='1':
                        masterMentorDict[row[2]]['merge_fields']['MMERGE13']=row[9]+" "+row[10]
                        masterMentorDict[row[2]]['merge_fields']['MMERGE14']=row[8]
                        masterMentorDict[row[2]]['merge_fields']['MMERGE19']=row[11]
                        masterMentorDict[row[2]]['count'] = 2
                    else:
                        masterMentorDict[row[2]]['merge_fields']['MMERGE16']=row[9]+" "+row[10]
                        masterMentorDict[row[2]]['merge_fields']['MMERGE17']=row[8]
                        masterMentorDict[row[2]]['merge_fields']['MMERGE20']=row[11]

                #mentor not found in master dict and adding for first time
                else:
                    mentorData = "{\n \"email_address\": \""+row[2]+"\", \"status\": \"subscribed\",\"count\": \"1\", \"merge_fields\": {\"FNAME\": \"" + row[3] + "\", \"MMERGE6\": \"" + row[9]+" "+row[10] + "\", \"MMERGE7\": \"" + row[8] + "\", \"MMERGE8\": \"" + row[11] + "\", \"MMERGE9\": \"Mentor\",\"MMERGE12\": \"" + row[6] + "\"}\n}"
                    masterMentorDict[row[2]] = json.loads(mentorData)



                #add todays timestamp for match date
                row[12] = today
                #append match to final match array
                finalMatchArray.append(row)
                menteeEmailArray.append(row[8])

    if sendMentorEmailBool == 'TRUE':
        for key in masterMentorDict.keys():
            row = masterMentorDict[key]
            sendMentorEmail = send_email(key,json.dumps(row),'PUT')
            
            print(sendMentorEmail.status_code)
            print(row)
            #email not found
            if sendMentorEmail.status_code == 404:
                sendMentorEmail.status_code = send_email(key,json.dumps(row),'POST')
    

            for idx, value in enumerate(potentialMatchValues):

                #find the email in the potenial match 
                if value[2]==key:
                    #update mentoremail status in potential matches spreadsheet
                    body = {'values': [[sendMentorEmail.status_code]]}
                    result = set_sheet('1QVPlpdxvGyQxB8jf6zD_Ci9xro-aXxVbyiwkRb6Oito','Sheet1!N'+str(idx+2),body)
                    if not result:
                        print('ERROR update mentor with email status code. Email: '+value[2]+' Status Code: '+sendMentorEmail.status_code)


 

 #get mentee waitlist to write timestamp of when matched
    menteeWaitlistValues = get_sheet('1vxqr3gdO-etFuOI7FV1HlVPr1q4r9fS0Mra3ExX_L6w','Sheet1!A2:V')
    if not menteeWaitlistValues:
        print('No mentee waitlist data returned.')

    for idx, row in enumerate(menteeWaitlistValues):

        #find the email in the final match array
        if row[3] in menteeEmailArray:
  
            #update mentee in waitlist
            body = {'values':[[datetime.now().strftime("%m/%d/%Y %H:%M:%S")]]}
            result = set_sheet('1vxqr3gdO-etFuOI7FV1HlVPr1q4r9fS0Mra3ExX_L6w','Sheet1!V'+str(idx+2),body)
            if not result:
                print('ERROR update mentee waitlist with timestamp. Email: '+row[3])


    #append the new matches to master list
    body = {'values':finalMatchArray}
    result = update_sheet('1Xna1oDmlBbH1GR6bItb9LwSQqFVHTKjNwVUYXkS6flA','Sheet1!A2:L',body)
    if not result:
       print('Final matches were not appended')
    else:
        print('Final matches have been sent and recorded')
   

if __name__ == '__main__':
    main()
