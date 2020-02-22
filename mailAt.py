import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase 
from email import encoders 

sender = input(str("Mail adresinizi giriniz: "))
password = input(str("Şifrenizi giriniz: "))

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(sender, password)
print("login success")


import pandas as pd


df = pd.read_excel (r'deneme.xlsx')
sayi = len(df)

mesaj = """mesajiniz"""

filename = "yourFile.pdf"
 # We assume that the file is in the directory where you run your Python script from
with open(filename, "rb") as attachment:
    # The content type "application/octet-stream" means that a MIME attachment is a binary file
    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())

# Encode to base64
encoders.encode_base64(part)

# Add header 
part.add_header(
    "Content-Disposition",
    f"attachment; filename= {filename}",
)



for x in range(sayi):
    kisi = df.loc[x]
    receiver = kisi["Mail"]
    cinsiyet = kisi["Cinsiyet"]
    ad = kisi["Ad"]
    if cinsiyet == 'f':
        hitap = "Merhaba " + ad + " Hanım,\n"
    elif cinsiyet == 'm':
        hitap = "Merhaba " + ad + " Bey,\n"
    gonder = hitap + mesaj
    gonder_text = MIMEText(gonder,"plain")
    message = MIMEMultipart()
    message["Subject"] = "Konu"
    message["From"] = sender
    message["To"] = receiver
    message.attach(gonder_text)
    message.attach(part)
    server.sendmail(sender, receiver, message.as_string())
    print("mail gonderildi")
