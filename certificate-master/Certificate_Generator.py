import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os

data = pd.read_excel (r'names.xlsx') 
name_list = data["Name"].tolist() 
for i in name_list:
    im = Image.open(r'template.png')
    d = ImageDraw.Draw(im)
    location = (100, 398)
    text_color = (0, 137, 209)
    font = ImageFont.truetype("arial.ttf", 70)
    d.text(location, i, fill = text_color, font = font)
    im.save("Certificate_" + i + ".pdf")
    os.startfile("Certificate_" + i + ".pdf")
