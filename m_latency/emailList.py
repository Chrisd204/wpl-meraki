import login
import requests, json
import datetime, sys, os, smtplib
import pandas as pd
import numpy as np
from pandas import ExcelWriter
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase 
from email import encoders


class SendEmail:
    def my_email(data):
        today = datetime.date.today()
        pd.options.display.float_format = '{:.1f}ms'.format

        msg = MIMEMultipart() 
        msg['From'] =login.monitor_email
        msg['To'] = login.my_email 
        msg['Subject'] = "Alert for Community Options Inc -All Mx's - Uplink Packet Loss & Latency"
        body = "Attached are updates for sites experiencing 100% packet loss for previous day, along with site latency averages."
        msg.attach(MIMEText(body, 'plain')) 
        filename = 'averages-'+str(today)+'.xlsx'
        attachment = open('/home/cdurham/Documents/code/wpl-meraki/averages-'+str(today)+'.xlsx', "rb") 
        p = MIMEBase('application', 'octet-stream') 
        p.set_payload((attachment).read()) 
        encoders.encode_base64(p)   
        p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        msg.attach(p) 
        s = smtplib.SMTP(login.smtp_server,login.smtp_port) 
        s.starttls() 
        s.login(login.monitor_email, login.monitor_email_password)
        text = msg.as_string()
        s.sendmail(login.monitor_email, login.my_email, text) 
        s.quit()


    def meraki_email(data):
        today = datetime.date.today()
        msg = MIMEMultipart() 
        msg['From'] =login.monitor_email
        msg['To'] = login.company_email 
        msg['Subject'] = "Alert for Community Options Inc -All Mx's - Uplink Packet Loss & Latency"
        body = "Attached are updates for sites experiencing 100% packet loss for previous day, along with site latency averages."
        msg.attach(MIMEText(body, 'plain')) 
        filename = 'averages-'+str(today)+'.xlsx'
        attachment = open('/home/cdurham/Documents/code/wpl-meraki/averages-'+str(today)+'.xlsx', "rb") 
        p = MIMEBase('application', 'octet-stream') 
        p.set_payload((attachment).read()) 
        encoders.encode_base64(p)   
        p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        msg.attach(p) 
        s = smtplib.SMTP(login.smtp_server,login.smtp_port) 
        s.starttls() 
        s.login(login.monitor_email, login.monitor_email_password)
        text = msg.as_string()
        s.sendmail(login.monitor_email, login.company_email, text) 
        s.quit()

    def no_data_email(data):
        today = datetime.date.today()        
        msg = MIMEMultipart() 
        msg['From'] =login.monitor_email
        msg['To'] = login.my_email 
        msg['Subject'] = "Alert for Community Options Inc -All Mx's - Uplink Packet Loss & Latency"
        body = "Look man, you figured it out, but nothing to see here!! Check back tomorrow"
        msg.attach(MIMEText(body, 'plain')) 
        filename = 'averages-'+str(today)+'.xlsx'
        attachment = open('/home/cdurham/Documents/code/wpl-meraki/averages-'+str(today)+'.xlsx', "rb") 
        p = MIMEBase('application', 'octet-stream') 
        p.set_payload((attachment).read()) 
        encoders.encode_base64(p)   
        p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        msg.attach(p) 
        s = smtplib.SMTP(login.smtp_server,login.smtp_port) 
        s.starttls() 
        s.login(login.monitor_email, login.monitor_email_password)
        text = msg.as_string()
        s.sendmail(login.monitor_email, login.my_email, text) 
        s.quit()