#Referenced Shiyue (Shay) Cheng, shiychen@cisco.com github site for api calls
import requests, json
import datetime, sys, os, smtplib
import pandas as pd
from pandas import ExcelWriter
import login

#random function for getting network id
def get_network_name(network_id, networks):
    return [element for element in networks if network_id == element['id']][0]['name']
    
#------ if statement to authenicate and make api request
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

# ------- iterate over workbook for sites with packet loss
    sites = pd.read_excel(w, sheet_name=None)
    site_keys = sites.keys()

    results = []
    for office in site_keys:
        try:
            for loss in sites[office]['lossPercent'].where(sites[office]['lossPercent'] > 4.0).dropna():
                results.append(office)
        except KeyError:
            continue
    final_results = list(dict.fromkeys(results))

# ------ move files to archive folder
    os.system('mv ~/Documents/code/wpl-meraki/*.xlsx ~/Documents/code/wpl-t-archive/')

email_body_df = pd.DataFrame(final_results, columns=['Sites above 4%'])
print(email_body_df)