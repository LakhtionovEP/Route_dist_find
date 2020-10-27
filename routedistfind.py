import json
import requests
import time
import openpyxl
import sys
import os
import re

#Функция подстановки координат в API OpenStreetMap

def dist_find(lat1, long1, lat2, long2):
    response = requests.get(f'https://routing.openstreetmap.de/routed-car/route/v1/driving/{long1},{lat1};{long2},'
                            f'{lat2}?overview=false&geometries=polyline&steps=true')
    todos = json.loads(response.text)
    dist = int(todos['routes'][0]['distance']/1000)
    return dist

#Загрузка данных в формате Excel

wb = openpyxl.load_workbook('dist_input.xlsx')
sheet = wb.active
c, s, u, oc = 0, 0, 0, 0
sheet['G1'] = 'Distance'
i = 2
try:
    wb.save('dist_input.xlsx')
except PermissionError:
    print('Please close the resulting file' + ' dist_input.xlsx ' + 'before start')
    os.system('pause')
while sheet['A' + str(i)].value is not None:
    oc += 1
    i += 1
t = time.time()
for i in range(2, oc + 2):
    c += 1
    sys.stdout.write('\r%s' % c + ' from ' + str(oc) + ' records has been obtained ')
    sys.stdout.write(str(int((time.time() - t) // 60)) + ' minutes ' +
                     str(float('{:.1f}'.format((time.time() - t) % 60))) + ' seconds has been spent')
    sys.stdout.flush()
    x1 = re.sub(',', '.', str(sheet['B' + str(i)].value))
    y1 = re.sub(',', '.', str(sheet['C' + str(i)].value))
    x2 = re.sub(',', '.', str(sheet['E' + str(i)].value))
    y2 = re.sub(',', '.', str(sheet['F' + str(i)].value))
    try:
        sheet['G' + str(i)].value = dist_find(x1, y1, x2, y2)
        s += 1
    except:
        sheet['G' + str(i)].value = 'Can\'t lay route for auto'
        u += 1
    time.sleep(1)

#Подсчет статистики

print()
print('Programme has finished')
print('Total records quantity: ' + str(c))
print('Successfull requests: ' + str(s))
print('Failed requests: ' + str(u))
print('Overall time: ' + (str(int((time.time() - t) // 60)) + ' minutes ' +
                          str(float('{:.1f}'.format((time.time() - t) % 60)))) + ' seconds')
wb.save('dist_input.xlsx')
os.system('pause')
