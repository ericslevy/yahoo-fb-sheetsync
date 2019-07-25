# yahoo-fb-sheetsync

This script updates a Google Sheet with data scraped from the Yahoo Fantasy Baseball website to create a league overview. Initially, this was done as a manual process whenever people got around to it.

## Getting Started

NOTE:  This script was designed to meet the needs of our specific league and will need to be edited further. I would like to add in some more customization down the line. Any number of teams is fine, but different positions(CI/MI/LF/CF/RF) would need to be fixed.  Colors in Google Sheets use the RGB 1.0 scale and would need to be edited manually throughout the code.

Our league uses minor league players, which are not stored in Yahoo. As they are a yearly selection process, they are kept on a hidden sheet and updated as needed. Lastly, we chose to use one of these minor leaguers to take up a space on the roster in order to allow owners with Pitcher/Batter to carry both sides without messing with roster caps for everyone else.  They are filtered out of the spreadsheet from Yahoo.

This script requires a few enviromental variables.  Since this is designed for Heroku, it utilizes os.environ.get() to grab config_vars set up in Heroku's settings. This would need to be edited for any other host as well.

Currently, the variables needed are:

dynasty_secret: A JSON [Google Sheets API OAuth token](https://developers.google.com/sheets/api/guides/authorizing)
league_id: a numeric league id obtained from the URL of any league page
oauth2_file: A JSON [Yahoo Developer Network API Oauth token](https://developer.yahoo.com/apps/)

## Authors

Eric Levy

## License

This project is licensed under the MIT License

