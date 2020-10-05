import login
import meraki_packetloss
from emailList import SendEmail
import requests, json
import datetime, sys, os, smtplib
import pandas as pd
import numpy as np
from pandas import ExcelWriter
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase 
from email import encoders



today = datetime.date.today()
w = pd.ExcelWriter('wpl-'+str(today)+'.xlsx')

def get_latency(file_name):

    global email_body_df # needed to use varibale outside of function.

    sites = pd.read_excel(w, sheet_name=None, dtype={'latencyMs':float})
    site_keys = sites.keys()

    results = []
    average_latencyMs = []
    average_lossPercent = []

    latencyMs_threshold = 150

    for office in site_keys:
        try:
            for latency in sites[office]['latencyMs'].dropna():
                latencyMs = sites[office]['latencyMs']
                average_for_latencyMs = latencyMs.mean()
                count = (sites[office]['latencyMs'] >= latencyMs_threshold).sum()
                if average_for_latencyMs >= 150 or count >= 10:
                    results.append(office)
                    average_latencyMs.append(average_for_latencyMs)

        except KeyError:
            continue

        list_results = list(dict.fromkeys(results))
        list_average_latencyMs = list(dict.fromkeys(average_latencyMs))
        
        s1 = pd.Series(list_results, dtype='string', name='Sites')
        s3 = pd.Series(list_average_latencyMs, dtype='float64', name='Latency in Ms')

        format_dict = {'latencyMs': '{:.2%}'}

        d = pd.concat([s1,s3], axis=1)
        email_body_df = pd.DataFrame(d)
      

def send_to_excel(df):
    global writer

    df = email_body_df
    writer = pd.ExcelWriter('latency-'+str(today)+'.xlsx')
    df.to_excel(writer, sheet_name='sites',float_format="%.1f")
    writer.save()

get_latency(w)

send_to_excel(email_body_df)

if email_body_df.empty:
    print ("Look man, you figured it out, but nothing to see here!! Check back tomorrow")
    SendEmail.latency_no_data_email(email_body_df)

else:
    print("Look's like we have big things to report,I'm going to send off the email! :)")
    SendEmail.latency_my_email(email_body_df)