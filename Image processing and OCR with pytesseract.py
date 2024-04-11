import glob
from PIL import Image, ImageEnhance
import os
import pandas as pd
import pytesseract

def pre_process(image_path, binarization_shreshold):
    image = Image.open(image_path)
    new_folder = image_path.strip('.'+image_path.split('.')[-1])
    os.makedirs(new_folder, exist_ok='True')

    enhancer = ImageEnhance.Brightness(image)
    image = enhancer.enhance(1.1)
    image.save(new_folder+'/1_Brightness enhanced image.png')

    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(1.1)
    image.save(new_folder+'/2_Contrast enhanced image.png')

    enhancer = ImageEnhance.Sharpness(image)
    image = enhancer.enhance(1.1)
    image.save(new_folder+'/3_Sharpness enhanced image.png')

    image = image.convert("L")
    image.save(new_folder+'/4_Greyscale image.png')

    table = []
    for i  in range(256):
        if i < binarization_shreshold:
            table.append(0)
        else:
            table.append(1)
    image = image.point(table,"1")
    image.save(new_folder+'/5_Binarized image.png')

    return image

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

script_path = os.path.abspath(__file__)
script_directory = os.path.dirname(script_path)
path = script_directory+'/Pictures\*.*'
filenames = []
image_formats = []
texts = []
for file in glob.glob(path):
    filename = file.split('\\')[-1].split('.')[0]
    image_format = file.split('\\')[-1].split('.')[1]
    try:
        image = pre_process(file, 200)
        text = pytesseract.image_to_string(image, lang='eng')
        filenames.append(filename)
        image_formats.append(image_format)
        texts.append(text)
        print(text)
    except:
        pass
    
df = pd.DataFrame()
df['Filename'] = filenames
df['Image Format'] = image_formats
df['Text'] = texts
df.to_excel(script_directory+'/Pictures/OCR_result.xlsx', index=False)