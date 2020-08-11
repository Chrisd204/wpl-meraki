#Referenced Shiyue (Shay) Cheng, shiychen@cisco.com github site for api calls
import requests, json
import datetime, sys, os, smtplib
import pandas as pd
from pandas import ExcelWriter
import login

#random function for getting network id
def get_network_name(network_id, networks):
    return [element for element in networks if network_id == element['id']][0]['name']

# ------- function iterate over workbook for sites with packet loss
def latency_averages(file_name):

    global email_body_df # needed to use varibale outside of function.

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
    w = pd.ExcelWriter('wpl-'+str(today)+'.xlsx') # creates file name
    for appliance in appliances:
            device_name = json.loads(session.get('https://api.meraki.com/api/v0/networks/' + appliance['networkId'] + '/devices/' + appliance['serial'], headers=headers).text)['name']
            packloss_latency = json.loads(session.get('https://api.meraki.com/api/v0/networks/'+appliance['networkId'] + '/devices/'+appliance['serial']+ '/lossAndLatencyHistory?uplink=wan1&ip=8.8.8.8&timespan=86400', headers=headers).text)
            try:
                print('Found appliance ' + device_name)
            except:
                print('Found appliance ' + appliance['serial'])
            df = pd.DataFrame(packloss_latency)
            df.to_excel(w, sheet_name=str(device_name), index=False)
            w.save()

# -- call function on excel workbook
    latency_averages(w)

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

    send_email(email_body_df)


    def send_to_excel(df):
        df = email_body_df
        writer = pd.ExcelWriter('averages-'+str(today)+'.xlsx')
        df.to_excel(writer, sheet_name='Averages')
        writer.save()
 
    send_to_excel(email_body_df)
