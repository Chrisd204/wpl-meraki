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
w = ('wpl-'+str(today)+'.xlsx')


def get_averages(file_name):
    global email_body_df # needed to use varibale outside of function.
    w_file = pd.ExcelFile(w)
    sheets = w_file.sheet_names

    results = []
    average_latencyMs = []
    average_lossPercent = []

    for site in sheets:
        org = w_file.parse(site)
        try:
            for loss in org['lossPercent']:
                if site not in results:
                    results.append(site)
                    lossPercent = org['lossPercent'].mean()
                    average_lossPercent.append(lossPercent)
                    latencyMs = org['latencyMs'].mean()
                    average_latencyMs.append(latencyMs)                    

        except KeyError:
            continue
    s1 = pd.Series(results,name='Sites')
    s2 = pd.Series(average_latencyMs,name='LatencyAverage')
    s3 = pd.Series(average_lossPercent,name='lossPercentAverage')        

    d = pd.concat([s1,s2,s3],axis=1)
    email_body_df = pd.DataFrame(d)

    print(len(results))
    print(len(average_lossPercent))
    print(email_body_df)


def send_to_excel(df):
    global writer

    df = email_body_df
    writer = pd.ExcelWriter('averages-'+str(today)+'.xlsx')
    df.to_excel(writer, sheet_name='sites',index=False,float_format="%.1f")
    writer.save()
get_averages(w)

send_to_excel(email_body_df)

#if email_body_df.empty:
#    print ("Look man, you figured it out, but nothing to see here!! Check back tomorrow")
#    SendEmail.latency_no_data_email(email_body_df)

#else:
#    print("Look's like we have big things to report,I'm going to send off the email! :)")
#    SendEmail.latency_my_email(email_body_df)

