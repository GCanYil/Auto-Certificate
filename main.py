import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os

if not os.path.exists("output"):
    os.makedirs("output")

df = pd.read_excel("data/veri.xlsx")
df["Tarih"] = df["Tarih"].dt.strftime("%d.%m.%Y")
rows_to_list = df.values.tolist()

name_font = ImageFont.truetype("assets/fonts/Mea_Culpa/MeaCulpa-Regular.ttf", 150)
description_font = ImageFont.truetype("assets/fonts/Roboto/Roboto-VariableFont_wdth,wght.ttf", 75)
date_font = ImageFont.truetype("assets/fonts/Roboto/Roboto-VariableFont_wdth,wght.ttf", 30)

for row in rows_to_list:
    name = row[0]
    type_of = "of achievement"
    certificate_date = row[1]
    email = row[2]

    img = Image.open("assets/certificate-templates/template1.png")
    width = img.width
    draw = ImageDraw.Draw(img)

    box = draw.textbbox((0,0), name, font=name_font)
    type_box = draw.textbbox((0,0), type_of, font=description_font)
    date_box = draw.textbbox((0,0), certificate_date, font=date_font)

    box_width = box[2] - box[0]
    type_box_width = type_box[2] - type_box[0]
    date_box_width = date_box[2] - date_box[0]

    x_coord = int((width - box_width) /2)
    type_box_x = int((width - type_box_width) /2)
    date_box_x = int((3400 - date_box_width) /2)

    draw.text((x_coord,600), name, font=name_font, fill=(0,0,0))
    draw.text((type_box_x,350), type_of, font=description_font, fill=(0,0,0))
    draw.text((date_box_x,1350), certificate_date, font=date_font, fill=(0,0,0))

    img.save(f"output/{name}.png")
