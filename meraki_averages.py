import login
import meraki_packetloss
from emailList import SendEmail
import requests, json
import datetime, sys, os, smtplib
import pandas as pd
import numpy as np
import statistics
from pandas import ExcelWriter
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase 
from email import encoders



today = datetime.date.today()
w = pd.ExcelWriter('wpl-'+str(today)+'.xlsx')

def get_averages(file_name):
    global email_body_df # needed to use varibale outside of function.
 
    sites = pd.read_excel(w,sheet_name=None, names=['startTs','endTs','lossPercent','latencyMs'])
    site_keys = sites.keys()

    results = []
    average_latencyMs = []
    average_lossPercent = []

    lossPercent_threshold = 100    

    for office in site_keys:
        try:
            for average in sites[office]['latencyMs']:
                results.append(office)
                latencyMs = sites[office]['latencyMs'].mean()
                average_latencyMs.append(latencyMs)
        except KeyError:
            continue



    list_results = list(dict.fromkeys(results))
    list_average_latencyMs = list(dict.fromkeys(average_latencyMs))
    list_average_lossPercent = list(dict.fromkeys(average_lossPercent))        

    s1 = pd.Series(list_results, dtype='string', name='Sites')
    s2 = pd.Series(list_average_latencyMs, dtype='float64', name='latencyMs')
    s3 = pd.Series(list_average_lossPercent, dtype='float64', name='lossPercent')        

    d = pd.concat([s1,s2,s3], axis=1)
    email_body_df = pd.DataFrame(d)

def send_to_excel(df):
    global writer

    df = email_body_df
    writer = pd.ExcelWriter('averages-'+str(today)+'.xlsx')
    df.to_excel(writer, sheet_name='sites',index=False,float_format="%.2f")
    writer.save()

get_averages(w)

send_to_excel(email_body_df)

#if email_body_df.empty:
#    print ("Look man, you figured it out, but nothing to see here!! Check back tomorrow")
#    SendEmail.latency_no_data_email(email_body_df)

#else:
#    print("Look's like we have big things to report,I'm going to send off the email! :)")
#    SendEmail.latency_my_email(email_body_df)

