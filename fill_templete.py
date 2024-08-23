import win32com.client
import pandas as pd
import os
import pyperclip
from utils import generate_qr_code, wrap_text

# подправить слово "Диаозон"  в русских шильдах. Разумеется на Диапозон.

# когда делаешь экспорт в аи появляется вот это окошко
# у тебя скорее всего стоит один из AI CS
# нужно чтобы было Adobe Illustrator 3JP

# При создании QR-кодов убирать ТИРЕ и ДЕФИСЫ в TAG номерах. Пример , в таблице : 2714-PG-469 . А QR-код должен быть 2714PG469.
# Таблички делаем только на то, на что есть серийный номер, таких 420 комплектов. Остальные позже как придут с завода изготовителя.

# распихать по папкам в новом порядке

LANG = 'eng'
LANG_RUS = 'инг' if LANG == 'eng' else 'рус'
SHIELD_WITHOUT_QR = 1
SHIELD_WITH_QR = 2
shield_type = SHIELD_WITH_QR
shield_type_text = "small" if shield_type == SHIELD_WITH_QR else "big"
MAX_TEXT_WIDTH = 4.330688976377953 if shield_type == SHIELD_WITHOUT_QR else 4.724389763779527
MAX_LINE_LENGTH = 50
SHOW_UI = shield_type == SHIELD_WITH_QR
ONLY_WITH_SERIAL_NUMBERS = True

# секундомер
import time
tempTime = time.time()
def getTime():
    global tempTime 
    offset = time.time() - tempTime
    tempTime = time.time()
    return offset
#####################################

def get_text(row, max_width):
    text = ''
    number_key = 'Код модели' if LANG == 'rus' else 'Model number'
    model_number = row[number_key]

    if shield_type == SHIELD_WITHOUT_QR:
        if LANG == 'eng':
            # Type 1 ENG
            text = f"Instrument tag no: {row['Instrument tag no']}\r" if row['Instrument tag no'] == row['Instrument tag no'] else f"Instrument tag no:"
            text += wrap_text(f"{number_key}: {model_number}", max_width) + "\r"
            text += f"Calibrated range: {row['Calibrated range']}\r" if row['Calibrated range'] == row['Calibrated range'] else f"Calibrated range:"
            text += f"Protective class: {row['Protective Class']}\r" if row['Protective Class'] == row['Protective Class'] else f"Protective class:"
            text += f"Climatic version: {row['Climatic version']}\r" if row['Climatic version'] == row['Climatic version'] else f"Climatic version:"
            text += f"Purchase order number: {row['Purchase order number']}\r" if row['Purchase order number'] == row['Purchase order number'] else f"Purchase order number"
            text += f"Serial number: {row['Serial number']}" if row['Serial number'] == row['Serial number'] else f"Serial number:"
        
        else:
            # Type 1 RUS
            text = f"Номер позиции: {row['Номер позиции']}\r" if row['Номер позиции'] == row['Номер позиции'] else f"Номер позиции:"
            text += wrap_text(f"{number_key}: {model_number}", max_width) + "\r"
            text += f"Диапазон измерения: {row['Диапазон измерения']}\r" if row['Диапазон измерения'] == row['Диапазон измерения'] else f"Диапазон измерения:"
            text += f"Степень защиты: {row['Степень защиты']}\r" if row['Степень защиты'] == row['Степень защиты'] else f"Степень защиты:"
            text += f"Климатическое исполнение: {row['Климатическое исполнение']}\r" if row['Климатическое исполнение'] == row['Климатическое исполнение'] else f"Климатическое исполнение:"
            text += f"Номер заказа: {row['Номер заказа']}\r" if row['Номер заказа'] == row['Номер заказа'] else f"Номер заказа"
            text += f"Серийный номер: {row['Серийный номер']}" if row['Серийный номер'] == row['Серийный номер'] else f"Серийный номер:"
    
    if shield_type == SHIELD_WITH_QR:
        if LANG == 'eng':
            # Type 2 ENG
            text = f"Instrument tag no: {row['Instrument tag no']}\r" if row['Instrument tag no'] == row['Instrument tag no'] else f"Instrument tag no:"
            text += wrap_text(f"Instrument service: {row['Instrument service']}", max_width) + '\r' if row['Instrument service'] == row['Instrument service'] else f"Instrument service:"
            text += wrap_text(f"Measured media: {row['Measured Media']}", max_width) + '\r' if row['Measured Media'] == row['Measured Media'] else f"Measured media:"
            text += f"Calibrated range: {row['Calibrated range']}\r" if row['Calibrated range'] == row['Calibrated range'] else f"Calibrated range:"
        
        else:
            # Type 2 RUS
            text = f"Номер позиции: {row['Номер позиции']}\r" if row['Номер позиции'] == row['Номер позиции'] else f"Номер позиции:"
            text += wrap_text(f"Функция: {row['Функция']}", max_width) + '\r' if row['Функция'] == row['Функция'] else f"Функция:"
            text += wrap_text(f"Измеряемая среда: {row['Измеряемая среда']}", max_width) + '\r' if row['Измеряемая среда'] == row['Измеряемая среда'] else f"Измеряемая среда:"
            text += f"Диапазон измерения: {row['Диапазон измерения']}\r" if row['Диапазон измерения'] == row['Диапазон измерения'] else f"Диапазон измерения:"
    
    
    return text

def process_template(cdr_file, xlsx_file, output_folder):
    data = pd.read_excel(xlsx_file)
    
    corel = win32com.client.gencache.EnsureDispatch("CorelDRAW.Application")
    corel.Visible = SHOW_UI
    
    getTime()
    # Главный цикл
    ##############################################################################
    for index, row in data.iterrows():
        tag_no = row['Instrument tag no']

        if row['Serial number'] == row['Serial number'] or not ONLY_WITH_SERIAL_NUMBERS:
            # try:
            #     src = open(cdr_file, 'rb')
            #     dst = open(os.path.join(output_folder, f"{tag_no}_{shield_type_text}_{LANG}.cdr"), 'wb')
            #     try:
            #         dst.write(src.read())
            #     finally:
            #         dst.close()
            # finally:
            #     src.close()

            # doc = corel.OpenDocument(os.path.join(output_folder, f"{tag_no}_{shield_type_text}_{LANG}.cdr"))
            doc = corel.OpenDocument(cdr_file)
            
            for page in doc.Pages:
                for shape in page.Shapes:
                    if shape.Type == 6:                        
                        max_line_length = MAX_LINE_LENGTH
                        shape.Text.Story.Text = get_text(row, max_line_length)

                        while shape.SizeWidth > MAX_TEXT_WIDTH:
                            if max_line_length < 20:
                                raise Exception('max_line_length = ' + str(max_line_length) + ' !')
                            max_line_length -= 1
                            shape.Text.Story.Text = get_text(row, max_line_length)

                        if shape.SizeWidth > MAX_TEXT_WIDTH:
                            raise Exception('shape.SizeWidth > MAX_TEXT_WIDTH !')
                        
                        shape.Fill.ApplyNoFill()
                        shape.Outline.Width = 0.003

            doc.ActivePage.Shapes.All().ConvertToCurves()
            
            if shield_type == SHIELD_WITH_QR:

                generate_qr_code(tag_no.replace('-',''))
                
                pyperclip.copy(os.path.join("D:\dev\shields\qr_codes", f"{tag_no.replace('-','')}.svg"))
                
                input("Ожидание импорта QR кода...")
            
                for page in doc.Pages:
                    page.Shapes.First.SetPosition(1.8300354330708661, 0.5905511811023622)
                    page.Shapes.First.SetSize(1.1811023622047243, 1.1811023622047243)
                    page.Shapes.First.Fill.ApplyNoFill()
                    page.Shapes.First.Outline.Width = 0.003

                pyperclip.copy(os.path.join(output_folder, f"{tag_no}_{shield_type_text}_{LANG}.ai"))
                input("Ожидание сохранения файла...")
            # Закрытие файла
            ##############################################################################
            # doc.Save()
            doc.Close()
            # os.remove(os.path.join(output_folder, f"Резервная_копия_{tag_no}_{shield_type_text}_{LANG}.cdr"))
            print(f"{str(index+1)}/{str(data.shape[0])}. Осталось {str((getTime() * (data.shape[0] - (index + 1)))/60)[0:5]} мин. Итерация: {str(getTime())[0:5]} сек..")

cdr_file =  f"D:\dev\shields\шильд тип {shield_type} {LANG_RUS}\шильд тип {shield_type} {LANG_RUS} ( текст еще не в кривых).cdr"
xlsx_file = "dataset.xlsx"
output_folder = f"D:\dev\shields\Type {shield_type}\{LANG.upper()}"

process_template(cdr_file, xlsx_file, output_folder)