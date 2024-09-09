# pywin32
import win32com.client
import pandas as pd
import os
import pyperclip
from utils import generate_qr_code, wrap_text

LANG = 'eng'
SHIELD_WITHOUT_QR = 1
SHIELD_WITH_QR = 2

MAX_LINE_LENGTH = 50

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

corel = win32com.client.Dispatch("CorelDRAW.Application")

def process_template(LANG, shield_type):
    LANG_RUS = 'инг' if LANG == 'eng' else 'рус'
    shield_type_text = "small" if shield_type == SHIELD_WITH_QR else "big"
    
    xlsx_file = "dataset.xlsx"
    folder = f"{os.getcwd()}\Type {shield_type}\{LANG.upper()}"

    data = pd.read_excel(xlsx_file)
    count = data.shape[0]
    if ONLY_WITH_SERIAL_NUMBERS:
        count = 0
        for index, row in data.iterrows():
            if row['Serial number'] == row['Serial number']:
                count += 1


    getTime()
    # Главный цикл
    ##############################################################################
    # doc = corel.OpenDocument(cdr_file)
    counter = 0
    for index, row in data.iterrows():

      tag_no = row['Instrument tag no']

      if (row['Serial number'] == row['Serial number'] or not ONLY_WITH_SERIAL_NUMBERS):
          counter += 1
          cdr_file = os.path.join(folder, f"{tag_no}_{shield_type_text}_{LANG}.ai")
          doc = corel.OpenDocument(cdr_file)
          # corel.Visible = False

          # for page in doc.Pages:
          #     for shape in page.Shapes:
          #       if shape.Type == 3:
          #         shape.Fill.UniformColor.RGBAssign(0, 0, 0)
          #         shape.Outline.Width = 0.003

          #         found = True

          # corel.Visible = True
          output_folder = f"{os.getcwd()}\PNG\Type {shield_type}\{LANG.upper()}"
          pyperclip.copy(os.path.join(output_folder, f"{tag_no}_{shield_type_text}_{LANG}.png"))
          input("Ожидание сохранения файла...")
          # Закрытие файла
          ##############################################################################
          doc.Close()
          interval = getTime()
          print(f"{str(counter)}/{str(count)}. Осталось {str((interval * (count - (counter)))/60)[0:5]} мин. Итерация: {str(interval)[0:5]} сек..")
          # print(row['Instrument tag no'])
          # print(row['Instrument service'])
          # print("index: "+ str(index*7))

corel.Visible = True
# process_template('rus', SHIELD_WITH_QR)
# process_template('eng', SHIELD_WITH_QR)

# process_template('rus', SHIELD_WITHOUT_QR)
process_template('eng', SHIELD_WITHOUT_QR)