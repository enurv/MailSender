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
#server.sendmail(sender, receiver, message)

import pandas as pd


df = pd.read_excel (r'deneme.xlsx')
#sayi = len(df)
sayi = 35

mesaj = """Sizinle İstanbul Üniversitesi-Cerrahpaşa Bilgisayar Kulübü adına iletişime geçiyoruz. Bilgisayar kulübü; okulumuz bilgisayar mühendisliği öğrencileri başta olmak üzere mühendislik fakültesinin bütün öğrencilerini bünyesinde barındıran, amacı mühendis adaylarını sektörle üniversite yıllarında buluşturmayı hedefleyen ve bu amaç  uğruna etkinlikler, eğitimler, projeler, teknik geziler ve çeşitli organizasyonlar düzenlemektir. Mühendisliğin sadece diploma almaktan ibaret olmadığını bilen bir kulüp olarak, mühendis adaylarını insanlığın faydasına dokunacak teknolojiler üretebilme becerisine sahip olmalarına katkıda bulunma isteğimiz var. İÜC Bilgisayar Kulübü, bu istek uğruna 2011’den beri faaliyet göstermektedir.

Size bu maili göndermemizin amacı iş birliği yapmak ve karşılıklı fayda sağlamaktır. Ekteki dosyalarda detaylı olarak anlatılan etkinliklerimizde yaklaşık 10.000 kişilik İstanbul Üniversitesi Cerrahpaşa Mühendislik Fakültesi’ne hitap etmekteyiz ve siz değerli markaları bu kitle ile buluşturmak için:

Bilişim sektörünü yakından tanıyabilmeleri için konferanslar ve atölyeler,
İş sahasını gözlemleyebilmek için teknik geziler,
Tam donanımlı mühendisler olabilmek ve iş dünyasında işe yarar yetenekler elde etmek için eğitimler,
“Mühendis üretir” mottosuyla yola çıkarak ekiplerin kulübümüz üyelerinden oluştuğu projeler,
Sadece öğrenmekle kalmayıp eğlenmeyi de unutmamak için çeşitli sosyal faaliyetler ve sosyal sorumluluk projeleri düzenlemekteyiz.

Bahsettiğimiz etkinliklerden en büyüğü olan ve gelenekselleşen “Bilişim Festivali” (BilFest) başta olmak üzere; etkinliklerimize ekteki sponsorluk dosyamızda belirtildiği üzere çeşitli yollarla destek olacak firmalarla iş birliği içerisine giriyoruz. Bu sene İÜC Bilgisayar Kulübü ailesinde sizleri aramızda görmeyi istiyoruz. 

Sizden ricamız bize verebileceğiniz bir randevuda sizlere hedeflerimiz ile çalışmalarımızı anlatma ve sizlerle tanışma imkânı sağlamaktır.
  
Olumlu ya da olumsuz dönüşünüz için teşekkür ederiz."""

filename = "Sponsorluk.pdf"
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
    message["Subject"] = "İÜ-CBK İşbirliği ve Sponsorluk"
    message["From"] = sender
    message["To"] = receiver
    message.attach(gonder_text)
    message.attach(part)
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(sender, password)
    server.sendmail(sender, receiver, message.as_string())
    print("mail gonderildi")