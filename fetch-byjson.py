from openpyxl import load_workbook
import requests
import json
import sys

wb = load_workbook('hostname-sheet-retrieve/inputhostnames.xlsx')
ws = wb.active
hostnames = []

for i in range(1, 201):
    cell = ws.cell(row = i, column = 1).value
    if cell:
        hostnames.append(cell)

wb.close()

url = 'https://jsonplaceholder.typicode.com/posts'

for j, hostname in enumerate(hostnames):
    with open('hostname-sheet-retrieve/body.json') as x:
        json_data = json.load(x)
        
        for data in json_data:
            data['id'] = hostname
            
    response = requests.post(url, json = json_data)
    file_path = 'hostname-sheet-retrieve/output/final{}.json'.format(j)
    
    with open(file_path, 'w') as outfile:
        if response.status_code == 201:
            json.dump(response.json(), outfile, indent = 4)
        else:
            print("Request failed")

sys.exit()