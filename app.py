
from __future__ import print_function
import os.path,time
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# https://drive.google.com/drive/folders/1hJVzRZoi6EJPp8Mn9n44GP8ZRxsu16eA
# https://docs.google.com/spreadsheets/d/1-qyAerPjUHYCxG0r_P0lbxphia3XqXMLzS4cCOmkhSU/edit#gid=1946276957

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/drive']


"""Shows basic usage of the Drive v3 API.
Prints the names and ids of the first 10 files the user has access to.
"""
creds = None
# The file token.json stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

service = build('drive', 'v3', credentials=creds)
Sheetservice = build('sheets', 'v4', credentials=creds)
sheet = Sheetservice.spreadsheets()


def GetExcelValues(range,Id):
    result = sheet.values().get(spreadsheetId=Id,
                                range=range).execute()
    values = result.get('values', [])
    if not values:
        print('error : Excel No data found.')
        return 0
    else:
        print('success: Excel Readable found')
        return values
def CheckFileDir(FileName,dir=1):
    # page_token = None
    if(dir == 1):

        results = service.files().list(q='mimeType = "application/vnd.google-apps.spreadsheet" and trashed=false',spaces='drive',fields="nextPageToken, files(id, name)",pageSize=400).execute()
    else:
        results = service.files().list(q='mimeType = "application/vnd.google-apps.folder" and trashed=false',spaces='drive',fields="nextPageToken, files(id, name)",pageSize=400).execute()
    items = results.get('files', [])

    # print(len(items))
    # for i in items:  
    if not items:
        print('No files found.')
        return None
    else:
        # print('Files:')
        for item in items:
            # print(item['name'])
            if(item['name'] == FileName):
                print(FileName + " is already there")
                # print(item['name'])
                return item['id']
def retrieve_permissions(file_id):
  """Retrieve a list of permissions.

  Args:
    service: Drive API service instance.
    file_id: ID of the file to retrieve permissions for.
  Returns:
    List of permissions.
  """
  try:
    permissions = service.permissions().list(fileId=file_id).execute()
    return permissions.get('permissions', [])
  except Exception as error:
    print('An error occurred: %s' % error)
  return None

def ShareFile(filename,emails):
    file_id = CheckFileDir(filename,0)
    print(file_id)
    perm_id = retrieve_permissions(file_id)
    print(perm_id)

    for id in perm_id:
        try:
            service.permissions().delete(fileId=file_id, permissionId=id['id']).execute()
        except Exception as e:
            print("Done deleting...")


    # add emails like this
    try:
        for email in emails:
            print("for email : "+email)
            new_permission = {
                'type': 'user',
                'role': 'reader',
                'emailAddress': email
                }
            try:
                run_new_permission = service.permissions().create(fileId=file_id,sendNotificationEmail=False,body=new_permission).execute()
                print("success : New Email added")
            except Exception as e:
                print("error: can't add permissions for " + email)
    except Exception as e:
        print("error : cant add new permission")



def GetExcelValues(range,Id):
    result = sheet.values().get(spreadsheetId=Id,
                                range=range).execute()
    # print(result)
    values = result.get('values', [])
    # print(values)
    if not values:
        print('error : Excel No data found.')
        return 0
    else:
        print('success: Excel Readable found')
        return values

if __name__ == '__main__':
    while True:    
        fileName = "Test"
        id = CheckFileDir(fileName)
        values = GetExcelValues('W9:W100000',id)
        emails = []
        for i in values:
            for email in i:
                emails.append(email)

        print(emails) 
        ShareFile('Foldertest',emails)
        print('info : sleeping for 10min 😊 ')
        time.sleep(600)