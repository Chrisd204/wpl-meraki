#Referenced Shiyue (Shay) Cheng, shiychen@cisco.com github site for api calls
import login
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

#random function for getting network id
def get_network_name(network_id, networks):
    return [element for element in networks if network_id == element['id']][0]['name']

# ------- function iterate over workbook for sites with packet loss
def get_packetloss(file_name):


    global email_body_df # needed to use varibale outside of function.

    sites = pd.read_excel(w, sheet_name=None, dtype={'latencyMs':float})
    site_keys = sites.keys()

    results = []
    average_latencyMs = []

    lossPercent_threshold = 100

    for office in site_keys:
        try:
            for loss in (sites[office]['lossPercent'] == lossPercent_threshold).dropna():
                count = (sites[office]['lossPercent'] == lossPercent_threshold).sum()
                if count >= 500 and count < 6000:
                    results.append(office)

                    latencyMs = sites[office]['latencyMs']
                    average_for_latencyMs = latencyMs.mean()
                    average_latencyMs.append(average_for_latencyMs)

        except KeyError:
            continue
        list_results = list(dict.fromkeys(results))
        list_average_latencyMs = list(dict.fromkeys(average_latencyMs))

        s1 = pd.Series(list_results, dtype='string', name='Sites')
        s2 = pd.Series(list_average_latencyMs, dtype='float64', name='Latency in Ms')

        d = pd.concat([s1,s2], axis=1)
        email_body_df = pd.DataFrame(d)



def send_to_excel(df):
    global writer

    df = email_body_df
    writer = pd.ExcelWriter('packetloss-'+str(today)+'.xlsx')
    df.to_excel(writer, sheet_name='sites',float_format="%.1f")
    writer.save()

if __name__ == '__main__':
    try:
        import login
        (API_KEY, ORG_ID) = (login.api_key, login.org_id)
    except ImportError:
        API_KEY = input('Enter your Dashboard API key: ')
        ORG_ID = input('Enter your organization ID: ')
    session = requests.session()
    headers = {'X-Cisco-Meraki-API-Key': API_KEY, 'Content-Type': 'application/json'} 
    try:
        name = json.loads(session.get('https://api.meraki.com/api/v0/organizations/' + login.org_id, headers=headers).text)['name'] # gets name of organization
    except:   
        sys.exit('Incorrect API key or org ID, as no valid data returned') # breaks if ID can't be found

#----- log in statement breaks, and now we create variables for applicances.
    networks = json.loads(session.get('https://api.meraki.com/api/v0/organizations/' + login.org_id + '/networks', headers=headers).text) # layer 1 grabs networks
    inventory = json.loads(session.get('https://api.meraki.com/api/v0/organizations/' + login.org_id + '/inventory', headers=headers).text) # layer 2 pulls full inventory
    appliances = [device for device in inventory if device['model'][:2] in ('MX') and device['networkId'] is not None] # layer 3 grabs only mx equipment
    devices = [device for device in inventory if device not in appliances and device['networkId'] is not None] # layer 4 sources everything else

#------------  Creates excel workbook with applicance data
    today = datetime.date.today()
    yesterday = today - datetime.timedelta(days=1)
    t0 = str(yesterday)+'T08:00:00Z'
    t1 = str(yesterday)+'T20:00:00Z'
    w = pd.ExcelWriter('wpl-'+str(today)+'.xlsx')
     # creates file name
    for appliance in appliances:

            device_name = json.loads(session.get('https://api.meraki.com/api/v0/networks/' + appliance['networkId'] + '/devices/' + appliance['serial'], headers=headers).text)['name']
            packloss_latency = json.loads(session.get('https://api.meraki.com/api/v0/networks/'+appliance['networkId'] + '/devices/'+appliance['serial']+ '/lossAndLatencyHistory?t0='+t0+'&t1='+t1+'&uplink=wan1&ip=8.8.8.8', headers=headers).text)
            try:
                print('Found appliance ' + device_name)
            except:
                print('Found appliance ' + appliance['serial'])
            df = pd.DataFrame(packloss_latency)
            df.to_excel(w, sheet_name=str(device_name),index=False,float_format="%.1f")
            w.save()


# -- call function on excel workbook
    get_packetloss(w)

# --- call function to send averages to excel
    send_to_excel(email_body_df)

# ------ send mail to company email with site list
    if email_body_df.empty:
        print ("Look man, you figured it out, but nothing to see here!! Check back tomorrow")
        SendEmail.no_data_email(email_body_df)

    else:
        print("Look's like we have big things to report,I'm going to send off the email! :)")
        SendEmail.my_email(email_body_df)