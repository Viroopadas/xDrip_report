import csv
import os
import sys
import re
from datetime import datetime, timedelta, time
import xlsxwriter
import PySimpleGUI as sg
import win32com.client as win32
from images import logo_path


folder_path = os.path.dirname(sys.argv[0])

def create_report(file_name: str, xe_type:int):
    # Преобразуем файл в словарь
    data = []
    try:
        with open(file_name, newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f, delimiter=';')
            for row in reader:
                data.append(row)
    except:
        window['_OUTPUT_'].update('')
        sg.popup('Файл CSV с таким именем не найден.', icon=logo_path)
        return False
            
    with open(file_name, newline='', encoding='utf-8') as f:
        try:
            reader = csv.DictReader(f, delimiter=';')
            for field in ['DAY', 'TIME', 'UDT_CGMS', 'BG_LEVEL', 'CH_GR', 'BOLUS', 'REMARK']:
                if field not in reader.fieldnames:
                    sg.popup('Неверный формат данных. Выберите другой файл.', icon=logo_path)
                    return False
            for row in reader:
                data.append(row)
        except:
            sg.popup('Неверный формат данных. Выберите другой файл.', icon=logo_path)
            return False

    blocks = {}

    # Создаем новую структуру данных из текущей
    for row in data:
        if row['BG_LEVEL']:
            continue

        date = datetime.strptime(row['DAY'], '%d.%m.%Y')
        glucose_level = f"{(float(row['UDT_CGMS']) / 18.02):.1f}" if row['UDT_CGMS'] else None
        carb = f"{(float(row['CH_GR']) / xe_types[xe_type]):.1f}" if row['CH_GR'] else None
        
        if date not in blocks:
            blocks[date] = []
                
        blocks[date].append({
            'time': row['TIME'],
            'glucose_level': glucose_level,
            'insulin': row['BOLUS'],
            'carb': carb,
            'remark': row['REMARK']
        })

    # Сортируем строки по времени и объединяем их в блоках по выборочным данным
    for date, block in blocks.items():
        block.sort(key=lambda x: x['time'])

        merged_block = []

        for row in block:
            if not merged_block or not (row['insulin'] or row['carb'] or row['remark']):
                merged_block.append(row)
            else:
                prev_row = merged_block[-1]
                if row['insulin']:
                    row['glucose_level'] = prev_row['glucose_level']          
                    merged_block.append(row)
                elif row['remark']:
                    prev_row['carb'] = prev_row['carb'] or row['carb']
                    prev_row['remark'] = prev_row['remark'] or row['remark']
                else:
                    merged_block.append(row)
        
        blocks[date] = merged_block

    # Создаем результирующую таблицу
    for date, block in blocks.items():
        result_block = []
        insulin_time = None
        amount_points = 0
        glucose_low = False
        glucose_high = False

        for row in block:
            current_time = datetime.combine(date, datetime.strptime(row['time'], '%H:%M').time())
            glucose = float(row['glucose_level'] or 0)

            if current_time.time() < time(7, 0) or current_time.time() > time(21, 0):
                amount_points = 0

            if row['remark']:
                result_block.append(row)
            elif row['insulin']:
                result_block.append(row)
                amount_points = 1
                insulin_time = current_time
            elif (insulin_time and (current_time - insulin_time > timedelta(minutes=58) and amount_points == 1)):
                result_block.append(row)
                amount_points = 2
            elif (insulin_time and (current_time - insulin_time > timedelta(minutes=118) and amount_points == 2)):
                result_block.append(row)
                amount_points = 0
            elif glucose <= 3.9:
                if not glucose_low:
                    result_block.append(row)
                    glucose_low = True
            elif glucose > 3.9:
                glucose_low = False
            elif glucose >= 15:
                if not glucose_high:
                    result_block.append(row)
                    glucose_high = True
            elif glucose < 15:
                glucose_high = False    

        blocks[date] = result_block

    # Перебираем файлы и находим максимальное значение x
    x = 1
    pattern = re.compile(r'report_(\d+)\.xlsx')
    files = os.listdir()
    for file in files:
        match = pattern.match(file)
        if match:
            number = int(match.group(1))
            x = max(x, number + 1)

    # Создаем файл Excel
    workbook = xlsxwriter.Workbook(f'report_{x}.xlsx')
    date_format = workbook.add_format({'num_format': 'dd.mm.yyyy'})

    f_title = workbook.add_format({'border': True, 'align': 'center', 'bold': True, 'text_wrap': True, 'valign': 'vcenter'})
    f_value = workbook.add_format({'border': True, 'text_wrap': True, 'valign': 'vcenter'})
    f_value_center = workbook.add_format({'border': True, 'valign': 'vcenter', 'align': 'center'})
    f_value_yellow = workbook.add_format({'border': True, 'valign': 'vcenter', 'align': 'center', 'bg_color': '#F2DCDB'})

    col = 0

    for date, rows in blocks.items():
        if not col:
            ws = workbook.add_worksheet(str(date.day))
        
        # Задаем ширину столбцов
        for i, width in enumerate([6, 6, 6, 6, 20, 2]):
            ws.set_column(col+i, col+i, width)
        
        # Альбомная ориентация
        ws.set_landscape()
        # Ширина полей страницы при печати
        ws.set_margins(left=0.272, right=0.272, top=0.16, bottom=0.16)
        # Установка пустых колонтитулов
        ws.margin_header = 0
        ws.margin_footer = 0
        # Тип бумаги. 9 - это формат А4
        ws.set_paper(9)
        # Центрировать горизонтально
        ws.center_horizontally()

        ws.fit_to_pages(1, 1)

        row = 0
        ws.merge_range(row, col, row, col+4, f'Дата: {str(date)[:10]}', f_title)
        row += 1

        header = ['Время', 'Ур-нь ГК', 'Инсу- лин', xe_type, 'Примечание']
        for i, col_name in enumerate(header):
            ws.write(1, col+i, col_name, f_title)
        row += 1

        for row_data in rows:
            if row_data['glucose_level']:
                glucose = float(row_data['glucose_level'])
                style = f_value_yellow if glucose <= 3.9 or glucose >= 15 else f_value_center
            else:
                style = f_value_center

            ws.write(row, col, row_data['time'], style)
            ws.write(row, col+1, row_data['glucose_level'], style)
            ws.write(row, col+2, row_data['insulin'], f_title)
            ws.write(row, col+3, row_data['carb'], f_value_center)
            ws.write(row, col+4, row_data['remark'], f_value)
            
            row += 1
        
        col = 0 if col == 12 else col + 6

    # Сохраняем файл Excel
    workbook.close()

    # Конвертация файла Excel в файл PDF
    try:
        excel_file = os.path.abspath(f'report_{x}.xlsx')
        pdf_file = os.path.abspath(f'report_{x}.pdf')
        excel = win32.Dispatch('Excel.Application')
        wb = excel.Workbooks.Open(excel_file)
        wb.ExportAsFixedFormat(0, pdf_file)
        wb.Close()
        excel.Quit()
    except:
        window['_OUTPUT_'].update('')
        sg.popup('Ошибка при создании файла PDF. Обратитесь к разработчикам.', icon=logo_path)
        return False

    return True

    # if not os.path.exists("read_files"):
    #     os.makedirs("read_files")
    # shutil.move(file, "read_files/" + file)


# Находим файл для чтения
files = os.listdir()
file_name = ''
for file in files:
    if file.startswith("export"):
        file_name = os.path.join(os.getcwd(), file)
        break

sg.theme('LightGreen')

xe_types = {'Углеводы, г.': 1, '1 ХЕ = 10 У': 10, '1 ХЕ = 12 У': 12, '1 ХЕ = 15 У': 15}

layout = [[sg.Text('Выберите файл .CSV')], 
          [sg.InputText(file_name), sg.FileBrowse(file_types=(("CSV Files", "*.csv"),), initial_folder=folder_path)],
          [sg.Text('Формат отображения количества углеводов:'),
           sg.Combo(list(xe_types.keys()), default_value='Углеводы, г.', key='_XE_TYPE_')],
          [sg.Submit(), sg.Text('', key='_OUTPUT_')]]

window = sg.Window('xDrip+ (Создание отчета)', layout, icon=logo_path)

while True:
    event, values = window.read()
    if event in (None, 'Exit', 'Cancel'):
        break
    if not values[0]:
        sg.popup('Выберите файл .CSV', icon=logo_path)
    elif event == 'Submit':
        window['_OUTPUT_'].update('Ожидайте...', text_color='red')
        window.refresh()

        is_ready = create_report(file_name=values[0], xe_type=values['_XE_TYPE_'])

        if is_ready:
            window['_OUTPUT_'].update('Готово! Файлы PDF и EXCEL созданы в той же папке.', text_color='red')

window.close()

