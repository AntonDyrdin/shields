import win32com.client
import pandas as pd
import os
import pyperclip
import qrcode
import qrcode.image.svg
#X= 61,483 мм

lang = 'eng'
lang_rus = 'инг'
shield_type = 2
shield_type_text = "small" if shield_type == 2 else "big"

import time
tempTime = time.time()
def getTime():
    global tempTime 
    offset = time.time() - tempTime
    tempTime = time.time()
    return offset
#####################################

def run(cdr_file_folder, xlsx_file, output_folder):
    data = pd.read_excel(xlsx_file)
    
    corel = win32com.client.gencache.EnsureDispatch("CorelDRAW.Application")
    corel.Visible = True
    
    getTime()
    for index, row in data.iterrows():
        tag_no = row['Instrument tag no']

        if index > -1:
            doc = corel.OpenDocument(os.path.join(cdr_file_folder, f"{tag_no}_{shield_type_text}_{lang}.cdr"))
            pyperclip.copy(os.path.join(output_folder, f"{tag_no}_{shield_type_text}_{lang}.ai"))
            
            w = input("Ожидание...")
            
            doc.Close()
            print(index)
        
        print(f"осталось {str((getTime() * (data.shape[0] - (index + 1)))/60)[0:5]} мин.")

cdr_file_folder =  f"D:\dev\shields\Type {shield_type} {lang.upper()}"
xlsx_file = "dataset.xlsx"
output_folder = f"D:\dev\shields\AI\Type {shield_type} {lang.upper()} AI"

run(cdr_file_folder, xlsx_file, output_folder)