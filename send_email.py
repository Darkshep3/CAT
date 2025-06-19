import smtplib
from smtplib import SMTP_SSL
from email.message import EmailMessage

def email_send(Subject, From, To, Content):
    if (type(To)) == (type(None)):
        return
    msg = EmailMessage()
    msg.set_content(Content, subtype = 'html')
    msg['Subject'] = Subject
    msg['From'] = From
    msg['To'] = To
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login("uremail@gmail.com", "app password here")
        server.send_message(msg)
        server.quit()

