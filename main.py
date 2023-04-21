from openpyxl import load_workbook

book = load_workbook(filename='shedule.xlsx')
ws = book.active


def createShedule(book):
    sheet, group_cell, group_cab = 0, 0, 0
    group = input("Введите вашу группу: ").upper()
    group = group.split('-')
    for sheets in book:
        if ((group[0]) in str(sheets)) and (group[1] in str(sheets)):
            sheet = sheets
    for row in sheet:
        for cell in row:
            if '-'.join(group) == cell.value:
                group_cell = cell.column_letter
                group_cab = chr(ord(group_cell) - 1)
    time_current = 4
    lesson_id = 4
    for i in range(4, 64):
        day = sheet['A' + str(i)].value
        if day is not None:
            print(day)
            if day == 'среда':
                stop = '19:05-19:50'
            else:
                stop = '17:20-18:05'
            time = '09:00-09:45'
            while time != stop and time_current < 70:
                lesson = sheet[group_cell + str(lesson_id)].value
                cab = str(sheet[group_cab + str(lesson_id)].value)
                lesson_id += 1
                if lesson is not None:
                    lesson = " ".join(lesson.split())
                    print(time, end=' ')
                    print(lesson, end=' ')
                    if cab != 'None':
                        cab = " ".join(cab.split())
                        print(cab)
                    else:
                        print("Дистанционно")
                time_current += 1
                time = sheet['B' + str(time_current)].value


createShedule(book)
