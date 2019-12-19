__author__ = 'MaryamZakani'

import json
import string
import xlsxwriter

members = {}
lists = {}
customFields = {}
customFields_values = {}
cards = []
headers = ['id', 'name', 'list', 'members', 'dateLastActivity', 'desc', 'shortUrl']
digit_columns = []

# Reading data from json file
with open('file.json') as json_file:
    data = json.load(json_file)
    for p in data['lists']:
        lists[p['id']] = p['name']

    for p in data['members']:
        members[p['id']] = p['fullName']

    for p in data['customFields']:
        customFields[p['id']] = p['name']

        if p['type'] == 'list':
            customFields_values[p['id']] = {}
            for q in p['options']:
                customFields_values[p['id']][q['id']] = q['value']['text']

    for p in data['cards']:
        card = {}
        card['id'] = p['id']
        card['list'] = lists[p['idList']]

        card['dateLastActivity'] = p['dateLastActivity']
        card['desc'] = p['desc']
        card['name'] = p['name']
        card['shortUrl'] = p['shortUrl']
        card['members'] = [members[q] for q in p['idMembers']]

        for q in p['customFieldItems']:
            if customFields[q['idCustomField']] not in headers:
                headers.append(customFields[q['idCustomField']])
            if q['idCustomField'] in customFields_values:
                card[customFields[q['idCustomField']]] = customFields_values[q['idCustomField']][q['idValue']]
            else:
                if 'text' in q['value']:
                    card[customFields[q['idCustomField']]] = q['value']['text']
                elif 'checked' in q['value']:
                    card[customFields[q['idCustomField']]] = q['value']['checked']

        cards.append(card)

print(headers)

# selected_lists = lists
# select a list to export
selected_lists = {'5dabe049bb931a238a4be686': 'Export to Google sheets'}

print(lists)
# Writing data to Excel file
# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('cards.xlsx')
worksheet = workbook.add_worksheet()

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': 1})

print(headers)
# Write data headers.
for col, header in enumerate(headers):
    worksheet.write(0, col, header, bold)

# Start from the first cell below the headers.
row = 1
col = 0

for item in cards:
    if item['list'] not in selected_lists.values():
        continue
    for col, header in enumerate(headers):
        if header not in item.keys():
            worksheet.write_string(row, col, '')
            continue

        try:
            # Find digit columns
            if item[header].isdigit():
                worksheet.write_number(row, col, float(item[header]))
                digit_columns.append(header)
            else:
                worksheet.write_string(row, col, str(item[header]))
        except:
            worksheet.write_string(row, col, str(item[header]))
    row += 1

# Write a total using a formula.
worksheet.write(row, 0, 'Total', bold)

# Write totals
for header in digit_columns:
    index = headers.index(header)
    char = list(string.ascii_uppercase)[index]
    worksheet.write(row, index, '=SUM({}2:{}{})'.format(char, char, row))

workbook.close()
