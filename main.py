import pandas as pd
from bs4 import BeautifulSoup

# Открытие файла HTML и создание BS-объекта
with open("body.html", "r", encoding="utf-8") as f:
    soup = BeautifulSoup(f, 'lxml')

# Поиск таблицы со строками
table = soup.find('tbody')
lines = table.findAll('tr')

# Селекторы для парсинга имен, оценок, состояний и проверок
name_selector = {'class': 'onkcGd YVvGBb VBEdtc-Wvd9Cc zZN2Lb-Wvd9Cc WAGvmf asQXV'}
mark_selector = {'class': 'QRiHXd gRisWe'}
state_selector = {'class': 'dDKhVc YVvGBb'}
check_selector = {'class': 'SxMnzc'}

# Создание временного списка для хранения данных о студентах
temp_result = []
for line in lines[1:]:
    user = []
    name = line.find('a', name_selector)
    marks = line.findAll('div', mark_selector)
    # Добавление имени студента
    if name:
        user.append(name.text)
    # Парсинг оценок с учетом наличия/отсутствия состояния и проверки
    count = 0
    for mark in marks:
        state = mark.find('span', state_selector).text
        if state:
            user.append('0.5')
            count += 0.5
        else:
            check = mark.find('span', check_selector)
            if check.text == 'Пропущен срок сдачи' or check.text == 'Назначено':
                user.append('0')
            elif check.text == 'Сдано':
                user.append('1')
                count += 1
    # Добавление суммарного количества баллов
    user.append(count)
    temp_result.append(user)

# Создание списка для хранения данных о студентах и их оценках
result = []
for i in range(len(temp_result[0])):
    result.append([])
for arr in temp_result:
    for i, element in enumerate(arr):
        result[i].append(element)

# Создание DataFrame из результата и запись в файл Excel
df = pd.DataFrame(result).T
with pd.ExcelWriter('classroomgrades.xlsx') as writer:
    df.to_excel(writer, index=False)