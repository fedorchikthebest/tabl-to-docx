import email.mime.multipart
import email.mime.text
import smtplib
import email
import email.mime.application
import os


def send_mail(mail, header, filename):
    with open('password.txt') as f:
        passwd = f.read()
    msg = email.mime.multipart.MIMEMultipart()
    msg['Subject'] = header
    msg['From'] = 'at@cdtb.net.ru'
    msg['To'] = 'gimonchik@gimonchik.ru'

    # PDF attachment
    fp=open(filename,'rb')
    att = email.mime.application.MIMEApplication(fp.read(),_subtype="pdf")
    fp.close()
    att.add_header('Content-Disposition','attachment',filename=os.path.split(filename)[-1])
    msg.attach(att)

    s = smtplib.SMTP('mail.nic.ru')
    s.starttls()
    s.login('at@cdtb.net.ru', passwd)
    s.sendmail('at@cdtb.net.ru',[mail], msg.as_string())
    s.quit()
