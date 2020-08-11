import datetime, sys
import pandas as pd
from pandas import ExcelWriter
import smtplib

# 'wpl-2020-08-10.xlsx'
def latency_averages(file_name):
    sites = pd.read_excel(w, sheet_name=None, dtype={'latencyMs':int})
    site_keys = sites.keys()

    averages = []
    results = []
    for office in site_keys:
        try:
            for loss in sites[office]['lossPercent'].where(sites[office]['lossPercent'] > 4.0).dropna():
                latencyMs_column = sites[office]['latencyMs']
                average = latencyMs_column.mean().round(2)
                results.append(office)
                averages.append(average)
        except KeyError:
            continue

    final_results = list(dict.fromkeys(results))
    final_averages = list(dict.fromkeys(averages))

    d = {'Sites':final_results, 'Latency Averages':final_averages}

    email_body_df = pd.DataFrame(d)
    print(email_body_df)

latency_averages('wpl-2020-08-10.xlsx')

# ------ send mail to company email with site list
def send_email(data):

    smtpObj = smtplib.SMTP(login.smtp_server,login.smtp_port)
    smtpObj.ehlo()
    smtpObj.starttls()
    message = ("""Subject: Alert for Community Options Inc -All Mx's - Uplink Packet Loss & Latency\n
    Updates for sites experiencing packet loss above 4 percent with average latency from past 24 hours.\n\n
    """ + str(data))
    smtpObj.login(login.lab_email,login.lab_email_password)
    smtpObj.sendmail(login.lab_email,login.company_email, message)
    smtpObj.quit()

send_email(latency_averages)