import requests
import json
import datetime
import openpyxl

today = datetime.date.today()       # Get today's date in YYYY-MM-DD format

keys = []
keys_w_add_drop = []
current_row = 2


datanyze_url = 'https://api.datanyze.com/alerts/?email=florian.parzhuber@cdnetworks.com&token=25c368f5afa267a1b31870fe8cda3c09&date=' + str(today)

url = requests.get(datanyze_url)
jsons = json.loads(url.text)


111

### FIND NUMBER OF ADDS AND DROP FOR A PARTICULAR CDN ###
for key in jsons:
    keys.append(key)

for i in range(len(keys)):
    keys_w_add_drop.append(keys[i])
    try:
        number_adds = len(jsons[keys[i]]["1"])
        keys_w_add_drop.append(number_adds)
    except KeyError:
        number_adds = 0
        keys_w_add_drop.append(number_adds)

    try:
        number_drops = len(jsons[keys[i]]["2"])
        keys_w_add_drop.append(number_drops)
    except KeyError:
        number_drops = 0
        keys_w_add_drop.append(number_drops)


print (keys_w_add_drop)

### NOW WE WILL DO DE ACTUAL WRITING TO AN EXCEL SPREADSHEET ###
wb = openpyxl.Workbook()
ws = wb.active
ws.cell(row = 1, column = 2).value ="Domain"
ws.cell(row = 1, column = 3).value ="Alexa Rank"
ws.cell(row = 1, column = 4).value ="CDN Technology"
ws.cell(row = 1, column = 5).value ="Add / Drop"


print (len(keys_w_add_drop))

key_number = int(len(keys_w_add_drop)/3)+12

for i in range(0,key_number,3):

    for x in range(keys_w_add_drop[i+1]):
        ws.cell(row = len(ws['B'])+1, column = 2).value = jsons[keys_w_add_drop[i]]["1"][x]["domain"]
        ws.cell(row = len(ws['C']), column = 3).value = jsons[keys_w_add_drop[i]]["1"][x]["alexa_rank"]
        ws.cell(row = len(ws['D']), column = 4).value = keys_w_add_drop[i]
        ws.cell(row = len(ws['E']), column = 5).value = "Add"

    
    #if (keys_w_add_drop[i+1] != 0):
        #current_row += 1

for i in range(0,key_number,3):
    for z in range(keys_w_add_drop[i+2]):
        ws.cell(row = len(ws['B'])+1, column = 2).value = jsons[keys_w_add_drop[i]]["2"][z]["domain"]
        ws.cell(row = len(ws['C']), column = 3).value = jsons[keys_w_add_drop[i]]["2"][z]["alexa_rank"]
        ws.cell(row = len(ws['D']), column = 4).value = keys_w_add_drop[i]
        ws.cell(row = len(ws['E']), column = 5).value = "Drop"

   # if (keys_w_add_drop[i+2] != 0):
        #current_row += 1

print (len(ws['B']))

wb.save("C:/Users/User/Desktop/Daily Datanyze Export/" + str(today) + '.xlsx') # Save data to worksheet with today's date as file name











