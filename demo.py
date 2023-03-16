import os 
import json
import time
import calendar
import requests
import xlsxwriter

img_folder = 'img'

#create img folder
if not os.path.exists('img'):
    os.makedirs(img_folder)

#read Json data
with open('MOCK_DATA.json', 'r') as f:
    json_data = json.load(f)

#img download and store in img folder
for img in json_data:
    response = requests.get(img['avatar'])
    id = img['id']
    with open(f'{img_folder}/image{id}.jpg', 'wb') as f:
        f.write(response.content)

# Group the data by company name
groups = {}
for item in json_data:
    company = item['company_name']
    if company not in groups:
        groups[company] = []
    groups[company].append(item)

#xlsx File name using timestamp
file_name = time.gmtime()
xlsx_name = str(calendar.timegm(file_name))+'.xlsx'
workbook = xlsxwriter.Workbook(xlsx_name)
worksheet = workbook.add_worksheet()
worksheet.set_column('B:M', 17)

#Bold and Font Size of Row rang 0 to 1
user_data = workbook.add_format({'bold': True})
user_data.set_font_size(18)
user_data.set_align('center')
user_data.set_align('vcenter')
user_data.set_bg_color('#BBE3BC')
user_data.set_font_name('Arial')

#Bold and font size of row rang 2 to 3
cell_format = workbook.add_format({'bold': True})
cell_format.set_font_size(12)
cell_format.set_align('center')
cell_format.set_align('vcenter')
cell_format.set_bg_color('#FEF2CD')
cell_format.set_font_name('Arial')

#Haders style in excelsheet
worksheet.merge_range(0, 0, 1, 12, 'Users Data', user_data)
worksheet.merge_range(2, 0, 3, 0, 'SR.', cell_format)
worksheet.merge_range(2, 1, 3, 1, 'Avatar', cell_format)
worksheet.merge_range(2, 2, 3, 2, 'ID', cell_format)
worksheet.merge_range(2, 3, 3, 3, 'First Name', cell_format)
worksheet.merge_range(2, 4, 3, 4, 'Last Name', cell_format)
worksheet.merge_range(2, 5, 3, 5, 'Email', cell_format)
worksheet.merge_range(2, 6, 3, 6, 'Gender', cell_format)
worksheet.merge_range(2, 7, 3, 7, 'Company Name', cell_format)
worksheet.merge_range(2, 8, 3, 8, 'Job Title', cell_format)
worksheet.merge_range(2, 9, 3, 9, 'Skills', cell_format)
worksheet.merge_range(2, 10, 2, 12, 'Car', cell_format)
worksheet.write(3, 10, 'Make', cell_format)
worksheet.write(3, 11, 'Model', cell_format)
worksheet.write(3, 12, 'Year', cell_format)

i, j, row = 1, 4, 4
cell_width = 50
cell_height = 50
for company, items in groups.items():
    for data in items:
        worksheet.set_row(j, 45)
        worksheet.write(row, 0, i)
        id = data['id']
        worksheet.insert_image(row,  1, f'{img_folder}/image{id}.jpg', {'width': 50, 'height': 50})
        worksheet.write(row,  2, data['id'])
        worksheet.write(row,  3, data['first_name'])
        worksheet.write(row,  4, data['last_name'])
        worksheet.write(row,  5, data['email'])
        worksheet.write(row,  6, data['gender'])
        worksheet.write(row,  7, data['company_name'])
        worksheet.write(row,  8, data['job_title'])
        skills = ', '.join(data['skills'])
        worksheet.write(row,  9, skills)
        worksheet.write(row,  10, data['car']['make'])
        worksheet.write(row,  11, data['car']['model'])
        worksheet.write(row,  12, data['car']['year'])
        i +=1
        j +=1
        row +=1

workbook.close()
