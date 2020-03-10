from yahoo_oauth import OAuth2
import xmltodict
from lxml import etree as ET
import pandas as pd
import pygsheets
import string
from pygsheets.sheet import SheetAPIWrapper
from collections import defaultdict
import argparse
import time
import os
import json
import tempfile
from contextlib import contextmanager
from datetime import date, timedelta


@contextmanager
def tempinput(data):
    temp = tempfile.NamedTemporaryFile(delete=False, mode='wt')
    temp.write(data)
    temp.close()
    try:
        yield temp.name
    finally:
        os.unlink(temp.name)

class YahooFantasyAPI:
       
    def fetchLeague(self):
        league_id = os.environ.get('league_id')
        tomorrow = str(date.today() + timedelta(days=1))
        session = self.getSession()
        r = session.get(
        'https://fantasysports.yahooapis.com/fantasy/v2/league/mlb.l.' + league_id + '/teams/roster;date=' + tomorrow + '/players'
        )
        
        return r
        

    def getSession(self):
        data = os.environ.get('oauth2_file')
        with tempinput(data) as tempfilename:   
            oauth = OAuth2(None, None, from_file=tempfilename)

        if not oauth.token_is_valid():
            oauth.refresh_access_token()

        return oauth.session
       
def positionAppend(row):
    if row['Extra Positions'] != None:
        row['Player Name'] = row['Player Name'] + ' (' + row['Extra Positions'] + ')'
        
    return row['Player Name']
                
def convertToColumn(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def mergerule(sheetId, startcol, endcol):
    
    return ({
        "mergeCells": {
            "range": {
            "sheetId": sheetId,
            "startRowIndex": 0,
            "endRowIndex": 1,
            "startColumnIndex": startcol,
            "endColumnIndex": endcol,
            },
            "mergeType": "MERGE_ALL"
            },
            },
            {
        "mergeCells": {
            "range": {
            "sheetId": sheetId,
            "startRowIndex": 1,
            "endRowIndex": 2,
            "startColumnIndex": startcol,
            "endColumnIndex": endcol,
            },
            "mergeType": "MERGE_ALL"
            },    
    },)
    
def grayformat(sheetId, team_id, cols):

    startcol = ((team_id - 1) * 2)
    endcol = (((team_id - 1) * 2)+2)
    
    if team_id % 2 == 0:  
        return {  
            "repeatCell": {
            "range": {
              "sheetId": wks.id,
              "startRowIndex": 1,
              "endRowIndex": cols + 3,
              "startColumnIndex": startcol,
              "endColumnIndex": endcol,
            },
            "cell": {
              "userEnteredFormat": {
                "backgroundColor": {
                  "red": 1,
                  "green": 1,
                  "blue": 1
                },
                "verticalAlignment" : "TOP",
                "textFormat": {
                  "foregroundColor": {
                    "red": 0,
                    "green": 0,
                    "blue": 0
                  },
                  "fontSize": 10,
                },
              },
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat,verticalAlignment)"
            },
            }

    else:
        return {  
            "repeatCell": {
            "range": {
              "sheetId": wks.id,
              "startRowIndex": 1,
              "endRowIndex": cols + 3,
              "startColumnIndex": startcol,
              "endColumnIndex": endcol,
            },
            "cell": {
              "userEnteredFormat": {
                "backgroundColor": {
                  "red": .847,
                  "green": .847,
                  "blue": .847
                },
                "verticalAlignment" : "TOP",
                "textFormat": {
                  "foregroundColor": {
                    "red": 0,
                    "green": 0,
                    "blue": 0
                  },
                  "fontSize": 10,
                },
              },
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat,verticalAlignment)"
            },
            }

def dl_na_formatter(player, red, green, blue, dl_na_teampos_list):
    column = int(player[0]) * 2
    converted_column = convertToColumn(column)
    playerpos = player[6]
    playername = player[2]
    team = player[0]
    teampos = team + playerpos
    row = (position_column.index(playerpos) + 4)
    cell_label = converted_column + str(row)
    value = teamposcells[teampos]
    start_index = value.find(playername)
    next_start_index = start_index + len(playername) + 1
    if len(dl_na_json[teampos]) == 0:            
        if next_start_index < len(value):
            dl_na_json_object = {
                            "repeatCell": {
                                "range": {
                                  "sheetId": wks.id,
                                  "startRowIndex": row -1, 
                                  "endRowIndex": row, 
                                  "startColumnIndex": column - 1,
                                  "endColumnIndex": column
                                },
                              "cell": {
                                "textFormatRuns": [
                                    {"startIndex": start_index,
                                    "format": {
                                      "foregroundColor": {
                                        "red": red,
                                        "green": green,
                                        "blue": blue
                                    },
                                }
                            },
                            {
                                    "startIndex": next_start_index,
                                    "format": {
                                      "foregroundColor": {
                                        "red": 0,
                                        "green": 0,
                                        "blue": 0
                                    },
                                }
                            },
                                                          ]
                    },
                    "fields": "textFormatRuns"
                },
                }
      
        else:  
            dl_na_json_object = {
                            "repeatCell": {
                                "range": {
                                  "sheetId": wks.id,
                                  "startRowIndex": row -1, 
                                  "endRowIndex": row, 
                                  "startColumnIndex": column - 1,
                                  "endColumnIndex": column
                                },
                              "cell": {
                                "textFormatRuns": [
                                    {"startIndex": start_index,
                                    "format": {
                                      "foregroundColor": {
                                        "red": red,
                                        "green": green,
                                        "blue": blue
                                    },
                                }
                            }
                                                          ]
                    },
                    "fields": "textFormatRuns"
                },
                }
        
        dl_na_json[teampos].append(dl_na_json_object)
          
    else:
        if next_start_index < len(value):
            dl_na_json_object = ({"startIndex": start_index,
                                    "format": {
                                      "foregroundColor": {
                                        "red": red,
                                        "green": green,
                                        "blue": blue
                                    },
                                }
                            },
                            {
                                    "startIndex": next_start_index,
                                    "format": {
                                      "foregroundColor": {
                                        "red": 0,
                                        "green": 0,
                                        "blue": 0
                                    },
                                }
                            },)
    
        
      
        else:  
            dl_na_json_object = {"startIndex": start_index,
                                    "format": {
                                      "foregroundColor": {
                                        "red": red,
                                        "green": green,
                                        "blue": blue
                                    },
                                }
                       }    
        dl_na_json[teampos][0]['repeatCell']['cell']['textFormatRuns'].append(dl_na_json_object)
    dl_na_teampos.append(teampos)

def clear_formatter(column, row):        
    clear_json = {
                            "repeatCell": {
                                "range": {
                                  "sheetId": wks.id,
                                  "startRowIndex": row -1, 
                                  "endRowIndex": row, 
                                  "startColumnIndex": column - 1,
                                  "endColumnIndex": column
                                },
                              "cell": {
                                "textFormatRuns": [
                                    {"startIndex": 0,
                                    "format": {
                                      "foregroundColor": {
                                        "red": 0,
                                        "green": 0,
                                        "blue": 0
                                    },
                                }
                            }
                                                          ]
                    },
                    "fields": "textFormatRuns"
                },
                }
    return clear_json
    
parser = argparse.ArgumentParser()
parser.add_argument("-r", "--reset", action="store_true")
parser.add_argument('-m', '--useminors', action='store_true')
args = parser.parse_args()    
        

api = YahooFantasyAPI()
xml = api.fetchLeague()

parseddata = xmltodict.parse(xml.content)['fantasy_content']['league']['teams']

teams = parseddata['team']
rosterlist = []
teamlist = []
team_idlist = []
team_email_list = []

for team in teams:

    team_name = team['name']
    teamlist.append(team_name)
    team_id = team['team_id']
    team_idlist.append(team_id)   
    manager_email = team['managers']['manager']['email']
    team_email_list.append(manager_email)
    roster = team['roster']
    count = roster['players']['@count']
    allplayers = roster['players']
    players = allplayers['player']

    for player in players:  

        player_name = player['name']['full']
        player_last = player['name']['last']
        player_id = player['player_id']
        player_team = player['editorial_team_abbr']
        current_position = player['selected_position']['position']
        eligible_position_list = []
        if current_position in ['C','1B','2B','3B','SS','OF','SP', 'RP']:          
            eligible_position_list.append(current_position)
        for position in player['eligible_positions']['position']:                
            if (position != current_position and position not in ['BN', 'DL', 'NA']) or position == 'Util':  
                eligible_position_list.append(position)
                
                
        eligible_positions = '/'.join(eligible_position_list)
        complete_player = (team_id+','+team_name+','+player_name+','+player_last+','+player_team+','+current_position+','+eligible_positions)
        rosterlist.append(complete_player)
       
df = pd.DataFrame([sub.split(",") for sub in rosterlist], columns =['Team ID', 'Team Name', 'Player Name', 'Player Last', 'Team', 'Current Position', 'Eligible Positions',])

df['Eligible Positions'].replace(regex=True, to_replace=r'/Util', value='', inplace=True)
df['Eligible Positions'].replace(regex=True, to_replace=r'U/t/i/l', value='Util', inplace=True)
df['Eligible Positions'].replace(regex=True, to_replace=r'^P', value='SP/RP', inplace=True)
df['Eligible Positions'].replace(regex=True, to_replace=r'/P', value='', inplace=True)
df[['Base Position','Extra Positions']] = df['Eligible Positions'].str.split('/', expand=True, n=1)
df['Player Name'] = df.apply(positionAppend, axis=1)
df = df.sort_values(by=['Player Last'])
df = df.drop(['Eligible Positions'], axis=1)
playerlist = df.values.tolist()
   
teamfill = []

for team in teamlist:
    teamfill.append(team)
    teamfill.append('')

emailfill = []    

for email in team_email_list:
    emailfill.append(email)
    emailfill.append('')
    
 
    
with tempinput(os.environ.get('dynasty_secret')) as tempfile:
    gc = pygsheets.authorize(service_file=tempfile)  

sheet = gc.open(os.environ.get('google_sheet_name'))
wks = sheet.worksheet('title', 'Current Rosters')
if args.useminors:
    hidden_wks = sheet.worksheet('title', 'Hidden Sheet')


if args.reset:
    wks.clear(fields='*')
lastcol = convertToColumn(len(teamfill))
teamrange = 'A1:' + lastcol + '1'
emailrange= 'A2:' + lastcol + '2'
mergelist = []
for i in range(1, len(teamfill), 2):
   mergeblock = mergerule(wks.id, i-1, i + 1)
   mergelist.append(mergeblock)
wks.update_values(crange = teamrange, values = [teamfill])
wks.update_values(crange = emailrange, values = [emailfill])
gc.sheet.batch_update(sheet.id, mergelist)

if args.useminors:
    position_column = ['C','1B','2B','3B','SS','OF','Util','SP', 'RP','MiLB']
else:
    position_column = ['C','1B','2B','3B','SS','OF','Util','SP', 'RP']
    
position_column_list = [position_column]

if args.reset:
    header_json = ({              
                      "updateBorders": {
                        "range": {
                          "sheetId": wks.id,
                          "startRowIndex": 0,
                          "endRowIndex": 2,
                        },
                        "innerVertical": {
                          "style": "SOLID",
                          "color": {
                                "red": 0,
                                "green": 0,
                                "blue": 0,                       
                            }, 
                        },
                      },
                    },
                    {  
                      "repeatCell": {
                        "range": {
                          "sheetId": wks.id,
                          "startRowIndex": 0,
                          "endRowIndex": 1
                        },
                        "cell": {
                          "userEnteredFormat": {
                            "backgroundColor": {
                              "red": 1,
                              "green": 1,
                              "blue": 0
                            },
                            "horizontalAlignment" : "CENTER",
                            "textFormat": {
                              "foregroundColor": {
                                "red": 0,
                                "green": 0,
                                "blue": 0
                              },
                              "fontSize": 10,
                              "bold": True
                            },
                          },
                        },
                        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
                      },
                    })
                
                
    gc.sheet.batch_update(sheet.id, header_json)
    wks.unlink()
    for i in range(1, len(teamfill),2):
        col = convertToColumn(i)       
        pos_range = col + '4:' + col + str((3+len(position_column)))
        wks.update_values(crange = pos_range, values = position_column_list, extend=True, majordim = 'COLUMNS')
        #time.sleep(1.5)       
    wks.link()    
       
    graylist = []        
    for id in team_idlist:
        grayblock = grayformat(wks.id, int(id), len(position_column))
        graylist.append(grayblock)
    gc.sheet.batch_update(sheet.id, graylist)    
     
teamposcells = defaultdict(list)
dl_list = []
na_list = []

starter_list = []

updated_playerlist = []

if args.useminors:
    minor_league_fillins = hidden_wks.cell('A26').value
for player in playerlist:    
    playername = player[2]
    if playername not in minor_league_fillins:
        teamposcells[str(player[0] + player[6])].append(playername)        
        updated_playerlist.append(player)
        if player[5] not in ('BN', 'DL', 'NA'):
            starter_list.append(player)
        if player[5] == 'DL':
            dl_list.append(player)
        if player[5] == 'NA':
            na_list.append(player)

 
flatlist = [item for sublist in updated_playerlist for item in sublist]
wks.unlink()
for team in team_idlist:
    count = flatlist.count(team)
    col = convertToColumn(int(team)*2)
    wks.update_value(col + '3', 'MLB: ' + str(count) + '/25')
wks.link()    
    
for id in team_idlist:
    for position in position_column:
        teampos = str(id+position)
        if position is not 'MiLB': 
            content = teamposcells[teampos]
            if len(content) > 1:
                content = '\r\n'.join(content)                
            else:
                content = ''.join(content)
            teamposcells[teampos] = content


dl_na_json  = defaultdict(list)
dl_na_teampos = []

for player in dl_list:
    dl_na_formatter(player, 1, 0, 0, dl_na_teampos)


for player in na_list:
    dl_na_formatter(player, 0, 0, 1, dl_na_teampos)
    
wks.unlink()       
for id in team_idlist:
        for position in position_column:      
            column_id = int(id) * 2
            column = convertToColumn(column_id)
            row = (position_column.index(position) + 4)
            cell = column + str(row)         
            teampos = str(id+position)
            if position is not 'MiLB': 
                content = teamposcells[teampos]
                wks.update_value(cell, content)
                #time.sleep(1.1)
            if teampos in set(dl_na_teampos):
                json = dl_na_json[teampos]
                gc.sheet.batch_update(sheet.id, json)
                #time.sleep(1.1)
            else:       
                json = clear_formatter(column_id, row)
                gc.sheet.batch_update(sheet.id, json)
            if args.useminors:
                if position == 'MiLB':
                    mcol = convertToColumn(int(id))
                    mcell = mcol + str(2)           
                    content = hidden_wks.cell(mcell).value
                    #time.sleep(1)
                    wks.update_value(cell, content)  
wks.link()                    
        
fixed_resize_list = []
for id in team_idlist:
    startcol = (int(id)-1) * 2
    endcol = startcol + 1
    resize_request =  {
      "updateDimensionProperties": {
        "range": {
          "sheetId": wks.id,
          "dimension": "COLUMNS",
          "startIndex": startcol,
          "endIndex": endcol
        },
        "properties": {
          "pixelSize": 40
        },
        "fields": "pixelSize"
        }
       },
    fixed_resize_list.append(resize_request)

gc.sheet.batch_update(sheet.id,fixed_resize_list)
          
auto_resize_list = []
for id in team_idlist:
    startcol = ((int(id)-1) * 2) + 1
    endcol = startcol + 1
    resize_request =  {
      "autoResizeDimensions": {
        "dimensions": {
          "sheetId": wks.id,
          "dimension": "COLUMNS",
          "startIndex": startcol,
          "endIndex": endcol
        }
      }
    }
    auto_resize_list.append(resize_request)
gc.sheet.batch_update(sheet.id,auto_resize_list)
       


    
        
    
    

    
    
    
    
            
        
       

        

        
        

        
    
    



