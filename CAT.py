import send_email
from send_email import email_send
import datetime
from datetime import timedelta
import openpyxl
import random

#current dates
class rn:
    date = datetime.datetime.now()
    month = (str(date))[6:7]
    day = (str(date))[8:10]
ps = openpyxl.load_workbook("flash_news.xlsx")
sheet = ps['Sheet1']
  
#setting up email subject
From = 'darkshep1234@gmail.com'
To = [""]
gif_chill = "https://i.ibb.co/f26GqMV/fangif.gif"
gif_last = "https://i.ibb.co/nQr1bMP/stars-gif.gif"

with open("D:\\Allen\\Random\\E\\cat pics\\allcats.txt", "r") as catfile:
    gif_cat = catfile.readlines()
    
#Extracting From Excel
for row in range(2, sheet.max_row):
    recipient_name = f'{sheet['A'+ str(row)].value} {sheet['B'+ str(row)].value}'
    email = sheet['C'+ str(row)].value
    To = str(email)
    hp_date = sheet['H'+ str(row)].value
    if (type(hp_date)) == (type(None)):
        continue
    class hp:
        month = (str(hp_date))[6:7]
        day =(str(hp_date))[8:10]
        
#Email Content
    date = str(hp.month)+ '/' + str(hp.day)+ '/2024'        
    delta = hp_date - rn.date
    delta = delta + timedelta(days=1)
    Subject = recipient_name + "\'s " + date + " Hydroponic Reminder " + str(abs((6-delta.days)+1))
    Subject2 = recipient_name + "\'s " + date + " Hydroponic Reminder 1"
    Subject3 = recipient_name + "\'s " + date + " Final Hydroponic Reminder Letter :("
    with open("D:\\Allen\\Random\\E\\cat pics\\splashtext.txt", "r") as splashfile:
        splash_text = random.choice(splashfile.readlines())
    splash_text2 = ''
    gif_cat2 = random.choice (gif_cat)
    if gif_cat2 == "https://i.ibb.co/6g9JcrD/wrongcat.png" or gif_cat2 == "https://i.ibb.co/Btw4Fns/wrongcat2cuteee.gif":
        splash_text2 = "wrong cat oops"    
    Content = f"""
        <html>
        <body style="font-family:&quot;Cambria&quot;,Roboto,Helvetica,Arial,sans-serif;letter-spacing:0.1px;line-height:15px;color:rgb(0,0,0);padding-bottom:4px"><font size="3">
        Dearest {recipient_name},
        <p>Your Daily Hydroponics starts in <span style = background-color:rgb(255,234,0)>{delta.days} days!!!</span></p>
        <p> <span style = background-color:rgb(255,234,0)>Daily Random Fact: {splash_text}</p> </span>
        <p>https://docs.google.com/spreadsheets/d/1sjvD6L2bh4hin1T2pjqBd1Lqe7_BUuAqwCv4SlqXw68/edit#gid=1577741213</p>
        <p> {splash_text2} </p> 
        <p>cute cats alleviates stress!</p>
        <p>please hold your purrs for 9 seconds as the cat loads its nine lives...</p>
        <img src="{gif_cat2}" alt= "kitty failed to load" width="256"> </body> 
        <body style = "line-height:1px><font size ="2">
        <p> feel free to reply with random facts or cat gifs to add to the list </p>
        <p> make sure to mark email as not spam so it doesnt go to spam :) or reply to email with some random msg idk <p>
        </body>
        <html>
        """
    Content2 = f"""
        <html>
        <body style="font-family:&quot;Cambria&quot;,Roboto,Helvetica,Arial,sans-serif;letter-spacing:0.1px;line-height:15px;color:rgb(0,0,0);padding-bottom:4px"><font size="3">
        Dearest {recipient_name},
        <p>Your Daily Hydroponics is in <span style = background-color:rgb(255,234,0)>1 week!!!</p></span>
        <p> Daily Random Fact {splash_text}</p>        
        <p>https://docs.google.com/spreadsheets/d/1sjvD6L2bh4hin1T2pjqBd1Lqe7_BUuAqwCv4SlqXw68/edit#gid=1577741213</p>
        <p>In the meantime, chilling...</p>
        <img src="{gif_chill}" alt= "chillin">
        <p> feel free to reply with random facts or cat gifs to add to the list </p>
        <p> make sure to mark email as not spam so it doesnt go to spam :) or reply to email with some random msg idk <p>
        <body>
        <html>
        """
    Content3 = f"""
        <html>
        <body style="font-family:&quot;Cambria&quot;,Roboto,Helvetica,Arial,sans-serif;letter-spacing:0.1px;line-height:15px;color:rgb(0,0,0);padding-bottom:4px"><font size="3">
        <body>Beloved {recipient_name},
        <p>It is with a heavy heart that I compose this letter, knowing that it carries the weight of our impending farewell. The time has come for us to part ways, each venturing forth into the unknown paths that lie ahead.</p>
        <p>Our shared experiences have shaped me in ways I cannot fully express. From the simplest of moments to the grandest of adventures, your presence has been a constant source of comfort and camaraderie.</p>
        <p>Your presence has been a beacon of light in the labyrinth of life, guiding me through the darkest of nights with the warmth of your friendship. Like stars scattered across the night sky, our moments together twinkle in the vast expanse of my mind.</p>
        <p>As we bid adieu, I find solace in the memories we've created together. They will serve as beacons of light, guiding me through the darkness of our separation.</p>
        <img src="{gif_last}" alt= "bye...">
        <p>May the stars of destiny adorn our journey with the radiant hues of reunion, as we glide upon the celestial canvas of our entwined destinies. <p>
        <p>With heartfelt regards, </p>
        <p>Some Random Kuei Student</p>
        <p></p>
        <body>
        <html>
        """
#checks date time then sends
    if delta.days <= 0:
        pass
    elif delta.days == 1:
        email_send(Subject3, From, To, Content3)
    elif delta.days < 6:
        email_send(Subject, From, To, Content)
    elif delta.days == 7:
        email_send(Subject2, From, To, Content2)
