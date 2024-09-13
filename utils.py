import os
import qrcode
import qrcode.image.svg
# regex
import re

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

    return os.path.join(f"{os.getcwd()}\qr_codes", f"{data}.svg")

def wrap_text(text, max_width, strict_single_line=False, spaced_dashes_already_replaced=False):
    lines = []
    words = re.split(r'( +|-|/)', text)

    current_line = ""
    for i in range(0, len(words)):
        if len(current_line) + len(words[i]) > max_width:
            # следующее слово выйдет за границу
            lines.append(current_line.strip().replace('\n', '\r'))
            current_line = words[i]
        elif len(words) > i + 2 and words[i + 1] == '-' and len(current_line) + len(words[i]) + 1 + len(words[i + 1]) > max_width:
            if len(current_line) + len(words[i]) + 1 > max_width:
                lines.append(current_line.strip().replace('\n', '\r'))
                current_line = words[i]
            else:
                current_line += words[i] + '-'
                lines.append(current_line.strip().replace('\n', '\r'))
                current_line = ''
        elif current_line != '' or (words[i] != '-' and words[i] != ' '):
            current_line += words[i]

    if current_line:
        lines.append(current_line.strip().replace('\n', '\r'))
        
    # for line in lines:
        # print('"' +line+'"')
    # print('')
    
    if strict_single_line and len(lines) > 1:
        if spaced_dashes_already_replaced:
            print("Не удалось вместить текст: " + text)
        else:
            return wrap_text(text.replace(' - ', '-'), max_width, True, spaced_dashes_already_replaced=True)
        
    return "\r".join(lines)


def wipe():
  for i in os.listdir('./Type 2/ENG/'):
    os.remove('./Type 2/ENG/' + i)
  for i in os.listdir('./Type 2/RUS/'):
    os.remove('./Type 2/RUS/' + i)
  for i in os.listdir('./Type 1/ENG/'):
    os.remove('./Type 1/ENG/' + i)
  for i in os.listdir('./Type 1/RUS/'):
    os.remove('./Type 1/RUS/' + i)
  for i in os.listdir('./PNG/Type 2/ENG/'):
    os.remove('./PNG/Type 2/ENG/' + i)
  for i in os.listdir('./PNG/Type 2/RUS/'):
    os.remove('./PNG/Type 2/RUS/' + i)
  for i in os.listdir('./PNG/Type 1/ENG/'):
    os.remove('./PNG/Type 1/ENG/' + i)
  for i in os.listdir('./PNG/Type 1/RUS/'):
    os.remove('./PNG/Type 1/RUS/' + i)