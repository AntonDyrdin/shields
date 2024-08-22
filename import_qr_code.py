import win32com.client
import pandas as pd
import os
import pyperclip
import qrcode
import qrcode.image.svg

lang = 'eng'
lang_rus = 'инг'
shield_type = 2
shield_type_text = "small" if shield_type == 2 else "big"

# секундомер
import time
tempTime = time.time()
def getTime():
    global tempTime 
    offset = time.time() - tempTime
    tempTime = time.time()
    return offset
#####################################

def generate_qr_code(data):
        factory = qrcode.image.svg.SvgPathImage
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=0,
        )

        qr.add_data(data)
        qr.make(fit=True)

        img = qr.make_image(image_factory=factory)

        img.save(f"qr_codes/{data}.svg")

def process_template(cdr_file, xlsx_file, output_folder):
    data = pd.read_excel(xlsx_file)
    
    corel = win32com.client.gencache.EnsureDispatch("CorelDRAW.Application")
    corel.Visible = True
    
    getTime()
    for index, row in data.iterrows():
        tag_no = row['Instrument tag no']

        if index > -1:
            pyperclip.copy(os.path.join("D:\dev\shields\qr_codes", f"{tag_no}.svg"))
            doc = corel.OpenDocument(os.path.join(output_folder, f"{tag_no}_{shield_type_text}_{lang}.cdr"))
            
            generate_qr_code(tag_no)
            
            w = input("Ожидание импорта QR кода...")
            
            for page in doc.Pages:
                page.Shapes.First.SetPosition(1.8300354330708661, 0.5905511811023622)
                page.Shapes.First.SetSize(1.1811023622047243, 1.1811023622047243)
                page.Shapes.First.Fill.ApplyNoFill()
                page.Shapes.First.Outline.Width = 0.003
                        

            
            doc.Save()
            doc.Close()
            print(index)
        print(f"осталось {str((getTime() * (data.shape[0] - (index + 1)))/60)[0:5]} мин.")

cdr_file =  f"D:\dev\shields\шильд тип {shield_type} {lang_rus}\шильд тип {shield_type} {lang_rus} ( текст еще не в кривых).cdr"
xlsx_file = "dataset.xlsx"
output_folder = f"D:\dev\shields\Type {shield_type} {lang.upper()}"

process_template(cdr_file, xlsx_file, output_folder)