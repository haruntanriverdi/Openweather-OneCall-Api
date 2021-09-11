# -*- coding: utf-8 -*-
# Openweather One Call API

import base64, requests, json
import smtplib
import time
from datetime import datetime
import xlrd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from prettytable import PrettyTable
import html
import sys
import yaml

conf = yaml.load(open('conf/config.yml'),yaml.FullLoader)

# Mail mariables
sender_mail = conf['gmail-user']['email']
sender_mail_password = conf['gmail-user']['password']
receiver_mail = conf['receiver-mail']['my-mail']
smtp_server = "smtp.gmail.com"
smtp_port = 587

loc = ("coordlist.xlsx")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

sheet.cell_value(0, 0)

base_url = "https://api.openweathermap.org/data/2.5/onecall?lat={}&lon={}&exclude={}&appid={}"
api_key = conf['openweather']['api']

mail_list = []
okunmayan = 0

def create_html_message(sender, to, subject, message_text, cc=False):
    message = MIMEMultipart("alternative")
    # message = MIMEText(message_text, 'html')
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    if cc:
        message['cc'] = cc
    part1 = MIMEText(message_text, "plain")
    part2 = MIMEText(message_text, "html")
    message.attach(part1)
    message.attach(part2)
    raw_message = base64.urlsafe_b64encode(message.as_string().encode("utf-8"))
    return message


def send_mail(msg, subject=None):
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.ehlo()
    server.starttls()
    server.login(sender_mail, sender_mail_password)

    if subject:
        msg = 'Subject: {}\n\n{}'.format(subject, msg)
    server.sendmail(sender_mail, receiver_mail, msg)
    server.quit()
    # print(msg.as_string())

def get_wind_speed(lat,lon):
    response = requests.get(base_url.format(lat,lon,"minutely,hourly",api_key))
    data = json.loads(response.text)
    hatakodu = response.status_code
    return data, hatakodu

try:
    for i in range(1,sheet.nrows):

            time.sleep(1)
            list = sheet.row_values(i)
            json_dict, hatakodu = get_wind_speed(list[1], list[2])
            if hatakodu == 200:
                time_zone_offset = json_dict['timezone_offset']
                daily_len = len(json_dict['daily'])

                daily_dt = json_dict['daily'][daily_len - 4]['dt']
                daily_wind_speed = json_dict['daily'][daily_len - 4]['wind_speed']
                daily_weather = json_dict['daily'][daily_len - 4]['weather'][0]['description']
                daily_icon = json_dict['daily'][daily_len - 4]['weather'][0]['icon']

                daily_icon_html = """<img src="http://openweathermap.org/img/wn/{}.png" width="100%">""".format(daily_icon)

                daily_dt_local = datetime.utcfromtimestamp(daily_dt + time_zone_offset).strftime('%d/%m/%Y %H:%M')

                if -3 <= daily_wind_speed <= 3:
                    list.append(daily_wind_speed)
                    list.append(daily_weather)
                    list.append(daily_icon_html)
                    mail_list.append(list)
                    #print(mail_list)
            else:
                msgHtml = """<p> %s hatası</p><br><b>%s</b>""" % (hatakodu,json_dict)
                subject = "API Hatası"
                message = create_html_message(sender_mail, receiver_mail, subject, msgHtml)
                send_mail(message.as_string())
                sys.exit()


    tabular_fields = ["Santral", "Latitude", "Longitude","İl", "Yatırımcı", "Max Rüzgar Hızı", "Hava Durumu","Icon"]
    tabular_table = PrettyTable()
    tabular_table.field_names = tabular_fields
    for i in mail_list:
        tabular_table.add_row(i)

    list_html_table = tabular_table.get_html_string()

    if mail_list:
        msgHtml = """
            <html>
                <head>
                <style>
                    body {
                        font-family: sans-serif;
                    }
                    table, th, td {
                        border: 1px solid black;
                        border-collapse: collapse;
                    }
                    th, td {
                        padding: 5px;
                        text-align: left;    
                    }    
                </style>
                </head>
            <body>
            <p>Okunamayan santral sayısı: %s </p><br>
        
            %s
        
            </body>
            </html>
            """ % (okunmayan, list_html_table)

        msgHtml = html.unescape(msgHtml)




        subject = "{} Tarihinde Ruzgar Durumu 3m/s ve Altında Olan {} Saha".format(daily_dt_local, len(mail_list))
        message = create_html_message(sender_mail, receiver_mail, subject, msgHtml)

        send_mail(message.as_string())

    else:
        msgHtml = """
            <p> {} Tarihinde Ruzgar Durumu 3m/s ve Altında Olan Saha Bulunamamıştır</p>
                """.format(daily_dt_local)
        subject = "{} Tarihinde Ruzgar Durumu 3m/s ve Altında Olan Saha Bulunamamıştır".format(daily_dt_local)
        message = create_html_message(sender_mail, receiver_mail, subject, msgHtml)
        send_mail(message.as_string())

except Exception as e:
    msgHtml = """
    <p> %s hatası</p>
        """ % (e)
    subject = "Wind Speed Kod Hatası"
    message = create_html_message(sender_mail, receiver_mail, subject, msgHtml)
    send_mail(message.as_string())
