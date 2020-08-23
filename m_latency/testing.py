import login
import meraki_latency
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

def get_averages(file_name):


    global email_body_df # needed to use varibale outside of function.

    sites = pd.read_excel(w, sheet_name=None, dtype={'latencyMs':float})
    site_keys = sites.keys()

    results = []
    average_latencyMs = []
    average_lossPercent = []

    threshold = 101
    for office in site_keys:
        try:
            for loss in sites[office]['lossPercent'].truncate(before = 670, after =1270).where(sites[office]['lossPercent'].truncate(before = 670, after =1270) >= threshold).dropna():
                results.append(office)

                latencyMs = sites[office]['latencyMs']
                average_for_latencyMs = latencyMs.mean()
                average_latencyMs.append(average_for_latencyMs)

        except KeyError:
            continue

        list_results = list(dict.fromkeys(results))
        list_average_latencyMs = list(dict.fromkeys(average_latencyMs))
        
        s1 = pd.Series(list_results, dtype='string', name='Sites')
        s3 = pd.Series(list_average_latencyMs, dtype='float64', name='Latency in Ms')


        d = pd.concat([s1,s3], axis=1)
        email_body_df = pd.DataFrame(d)

    if email_body_df.empty:
        print("Look man, you figured it out, but nothing to see here!! Check back tomorrow")
    else:
        return print(email_body_df)
get_averages(w)

#.truncate(before = 670, after =1270)

#.truncate(before = 670, after =1270)