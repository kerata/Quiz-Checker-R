import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

color_default = {'red': 217/255, 'green': 217/255, 'blue': 217/255, 'alpha': 1}
color_wrong = {'red': 255/255, 'green': 217/255, 'blue': 102/255, 'alpha': 1}
color_header = {'red': 176/255, 'green': 179/255, 'blue': 178/255, 'alpha': 1}
color_id = {'red': 212/255, 'green': 212/255, 'blue': 212/255, 'alpha': 1}
color_white = {'red': 1.0, 'green': 1.0, 'blue': 1.0, 'alpha': 1}
color_black = {'red': 0.0, 'green': 0.0, 'blue': 0.0, 'alpha': 1}

border_solid = {"style": "SOLID", "color": color_black}

# TODO: should be updated.
spreadsheet_id = ''
section1_sheet_id = 0
section2_sheet_id = 0


def connect_GSheets():
    """Connects to the Sheets API.
    """

    scopes = 'https://www.googleapis.com/auth/spreadsheets'

    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        # if creds and creds.expired and creds.refresh_token:
        #     creds.refresh(Request())
        # else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', scopes)
        creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    return service


def get_border_request(cell_range):
    return {
      "updateBorders": {
        "range": cell_range,
        "top": border_solid,
        "bottom": border_solid,
        "left": border_solid,
        "right": border_solid,
        "innerHorizontal": border_solid,
        "innerVertical": border_solid
      }
    }
