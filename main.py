import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
import mailInfo
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import time
import qr_maker

sender_mail = mailInfo.mail
sender_pass = mailInfo.password

def send_mail(receiver_mail, receiver_name, certificate_path):
    fin = MIMEMultipart()
    fin["From"] = sender_mail
    fin["To"] = receiver_mail
    fin["Subject"] = "Congratulations! Your certificate is ready!"

    message_body = f"""Hey {receiver_name}!
We congratulate you for completing your course.
Your personal certificate is in the attachment.

Wishing you a great career!"""
    fin.attach(MIMEText(message_body, "plain"))

    try:
        with open(certificate_path, "rb") as certificate:
            certificate_data = certificate.read()
        image_attachment = MIMEImage(certificate_data, name = os.path.basename(certificate_path))
        fin.attach(image_attachment)
    except FileNotFoundError:
        print(f"Error: {certificate_path} couldn't be found. No mail sent.")
        return
    
    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_mail, sender_pass)
        message = fin.as_string()
        server.sendmail(sender_mail, receiver_mail, message)
        print(f"Certificate sent to -> {receiver_mail}")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        server.quit()

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

    qr_img = qr_maker.make_qr(name= name, date= certificate_date)
    qr_img = qr_img.resize((150,150))
    img.paste(qr_img, (1830,1240))

    img.save(f"output/{name}.png")
    send_mail(receiver_mail=email, receiver_name=name, certificate_path=f"output/{name}.png")
    time.sleep(2)
