# BlendSaver
A program that utilizes the Spotify and GSheets/GDrive APIs to save a spotify Blend playlist.
The Blend playlist updates once a day with 50 songs. This program archives the daily contents (song title, artist name, artist genres, popularity score, main listener) so you can keep track of the playlist's contents after previous iterations have expired.


## Setup
* Create a Spotify for Developers account and initialize the project, saving the credentials
* Create a Google Cloud project that contains a service account with access to the GSheets and GDrive API, replacing the credentials.json file with your new credentials
* Fill in the missing information in main.py
  * 'CLIENT_ID' on line 4
  * 'CLIENT_SECRET' on line 5
  * 'playlist_id' on line 15
  * 'folder_id' on line 299
* Run once to initialize and sign into Google
* Subsequent runs should not require sign in for up to 1 week

# Usage
After Setup, simply run the program once a day to archive the contents of your daily Blend playlist.
I recommend setting up a vbs and/or batch file to run the program, and schedule the task to run once a day.
You will be asked to sign in through google once a week, as your credentials expire after one week.