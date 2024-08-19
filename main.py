import spotipy
from spotipy.oauth2 import SpotifyOAuth

CLIENT_ID = '' # Add your client id from a Spotify dev account here
CLEINT_SECRET = '' # Add your client secret from a Spotify dev account here
REDIRECT_URI = 'https://google.com'

scope = 'playlist-read-private'

sp = spotipy.Spotify(auth_manager=SpotifyOAuth(client_id=CLIENT_ID,
                                               client_secret = CLEINT_SECRET, 
                                               redirect_uri=REDIRECT_URI, 
                                               scope=scope))

playlist_id = '' # Enter the playlist ID for your Blend playlist here

def getTracklist(): # Gathers all track data

    response = sp.playlist_tracks(playlist_id)
    tracks = response['tracks']
    num_songs = tracks['total']
    print(f"{num_songs} songs loaded")
    return tracks

def getGenreInfo(tracks): # Gathers genres of all involved artists
    print("Getting genres...")
    genre_list = [[]] * tracks['total']
    for i, item in enumerate(tracks['items']):
        genres = []
        for artist in item['track']['artists']:
            artist_id = artist['id']
            response = sp.artist(artist_id)
            new_genres = response['genres']
            genres += new_genres
        genre_list[i] = genres
    print("Genre list successfully generated")
    return genre_list



# -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=- #
#                                   #
#    Google Sheets/Drive Section    #
#                                   #
# -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=- #

from google.oauth2 import credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import datetime
import os.path
import pickle


SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']

def getDate():
    current_time = datetime.datetime.now()
    date = f"{current_time.month}/{current_time.day}/{current_time.year}"
    return date

def clearToken(): # Removes token file once a week to prevent credentials from expiring
    if datetime.datetime.today().weekday() == 5: # Saturday
        file_path = 'token.pickle'
        if os.path.exists(file_path):
            os.remove(file_path)
            print(f'{file_path} removed for renewal')
        else:
            print(f'{file_path} could not be found')


def authenticate():
    # Authenticate and return service objects
    creds = None
    # Token file to store user access and refresh tokens
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                './credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    drive_service = build('drive', 'v3', credentials=creds)
    sheets_service = build('sheets', 'v4', credentials=creds)
    return drive_service, sheets_service




def create_google_sheet(sheets_service, today):
    # Create a new spreadsheet
    sheet = {
        'properties': {
            'title': f'{today} Daily Blend'
        }
    }
    request = sheets_service.spreadsheets().create(body=sheet, fields='spreadsheetId')
    response = request.execute()
    spreadsheet_id = response.get('spreadsheetId')
    print(f"Created spreadsheet with ID: {spreadsheet_id}")
    return spreadsheet_id



def move_sheet_to_folder(drive_service, file_id, folder_id):
    # Move sheet to relevant folder

    # Get the existing permissions of the file
    file = drive_service.files().get(fileId=file_id, fields='parents').execute()
    previous_parents = ",".join([p for p in file.get('parents')])

    # Move the file to the new folder
    drive_service.files().update(
        fileId=file_id,
        addParents=folder_id,
        removeParents=previous_parents,
        fields='id, parents'
    ).execute()


def formatSheet(sheets_service, spreadsheet_id, tracks, genre_list):
    print("Formatting...")

    tracklist = []
    for i in range(tracks['total']):
        track = tracks['items'][i]['track']
        tracklist.append(track)

    requests = [ # Original Formatting
        {
            'repeatCell': {
                'range': {
                    'sheetId': 0,
                    'startRowIndex': 0,
                    'endRowIndex': 650,
                    'startColumnIndex': 0,
                    'endColumnIndex': 26,
                },
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColor': {
                            'red': 0.5,
                            'green': 0.93,
                            'blue': 1.0,
                        },
                        'textFormat': {
                            'fontSize': 14,
                            'fontFamily': 'courier new',
                        }
                    },
                },
                'fields': 'userEnteredFormat(backgroundColor,textFormat)',
            },
        },
    ]

    num_merges = len(tracklist)
    merge_height = 10
    break_size = 2
    song_height = merge_height + break_size

    for i in range(num_merges):
        start_row_index = i * song_height
        end_row_index = start_row_index + merge_height

        requests.append({ # Merge Cells for image
            'mergeCells': {
                'range': {
                    'sheetId': 0,
                    'startRowIndex': start_row_index,
                    'endRowIndex': end_row_index,
                    'startColumnIndex': 0,
                    'endColumnIndex': 2,
                },
                'mergeType': 'MERGE_ALL'
            }
        })

        requests.append({ # Merge cells for text
            'mergeCells': {
                'range': {
                    'sheetId': 0,
                    'startRowIndex': start_row_index,
                    'endRowIndex': end_row_index,
                    'startColumnIndex': 2,
                    'endColumnIndex': 26,
                },
                'mergeType': 'MERGE_ALL'
            }
        })



    start_row = 0
    start_column = 0
    for i, item in enumerate(tracklist): # Fill in all images
        row_index = start_row + (i * song_height)
        image_url = item['album']['images'][1]['url']
        requests.append({
            'updateCells': {
                'range': {
                    'sheetId': 0,
                    'startRowIndex': row_index,
                    'endRowIndex': row_index + 1,
                    'startColumnIndex': start_column,
                    'endColumnIndex': start_column + 1,
                },
                'rows': [{
                    'values': [{
                        'userEnteredValue': {
                            'formulaValue': f'=IMAGE("{image_url}")'
                        },
                        'userEnteredFormat': {
                            'horizontalAlignment': 'CENTER',
                            'verticalAlignment': 'MIDDLE',
                        },
                    }]
                }],
                'fields': 'userEnteredValue, userEnteredFormat(horizontalAlignment,verticalAlignment)'
            }
        })
        song_name = item['name']

        artist_list = item['artists'][0]['name']
        for artist in item['artists'][1:]:
            next_artist = artist['name']
            artist_list += f', {next_artist}'

        popularity = item['popularity']

        added_by = tracks['items'][i]['added_by']['id']

        if(genre_list[i] == []):
            genre_desc = 'Genre'
            genre = 'N/A'
        else:
            if(len(genre_list[i]) == 1):
                genre_desc = 'Genre'
            else:
                genre_desc = 'Genres'
            
            genre = genre_list[i][0]
            for item in genre_list[i][1:]:
                next_genre = item
                genre += f', {next_genre}'

        description = f'{song_name}\nby\n{artist_list}\n\n{genre_desc}: {genre}\n\nPopularity: {popularity}\n\nMain Listener: {added_by}'

        requests.append({
            'updateCells': {
                'range': {
                    'sheetId': 0,
                    'startRowIndex': row_index,
                    'endRowIndex': row_index + 1,
                    'startColumnIndex': 2,
                    'endColumnIndex': 3,
                },
                'rows': [{
                    'values': [{
                        'userEnteredValue': {
                            'stringValue': description
                        }
                    }]
                }],
                'fields': 'userEnteredValue'
            }
        })

    body = {
        'requests': requests
    }

    try:
        response = sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()
        print('Formatting applied')
    except HttpError as error:
        print(f'An error occurred: {error}')

def main():
    today = getDate()
    clearToken()
    tracks = getTracklist()
    genre = getGenreInfo(tracks)
    try:
        drive_service, sheets_service = authenticate()
        folder_id = ''  # Replace with the ID of the Google Drive folder that you wants to place the saved Blends into
        spreadsheet_id = create_google_sheet(sheets_service, today)
        move_sheet_to_folder(drive_service, spreadsheet_id, folder_id)
        print(f"Moved spreadsheet to folder with ID: {folder_id}")
    except HttpError as error:
        print(f'An error occurred: {error}')
        exit()
    
    formatSheet(sheets_service, spreadsheet_id, tracks, genre)

    print("Playlist successfully logged")
    


if __name__ == '__main__':
    main()