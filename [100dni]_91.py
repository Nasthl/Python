
import smtplib
from email.message import EmailMessage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email import encoders

msg= MIMEMultipart()
msg['From'] = 'tkdarkses@gmail.com'
msg['To'] = 't.korzyniec@gmail.com'
msg['Subject'] = 'Testowy mail'
tekst='testowy mail'
msg.attach(MIMEText(tekst))

with open('C:\\Users\\tkdar\Desktop\do skopiowania\zdjecia\\WP_20180318_14_47_59_Pro.jpg', 'rb') as f:
    img = MIMEImage(f.read(),name='C:\\Users\\tkdar\Desktop\do skopiowania\zdjecia\\WP_20180318_14_47_59_Pro.jpg')
    msg.attach(img)

#Konfiguracja serwera
mailserver = smtplib.SMTP('smtp.gmail.com',587)
mailserver.starttls()
mailserver.login('tkdarkses@gmail.com','xxx!')
mailserver.send_message(msg)
mailserver.quit()


