# packet loss measuring alerting on over 3% packet loss
import pandas as pd
import smtplib, ssl
import login

excel_file = 'wpl-2020-07-05.xlsx'

sites = pd.read_excel(excel_file, sheet_name=None)
site_keys = sites.keys()
results = []

for i in site_keys:
    try:
        for x in sites[i]['lossPercent'].where(sites[i]['lossPercent'] > 4.0).dropna():
            results.append(i)
    except KeyError:
        continue
final_results = list(dict.fromkeys(results))

print(final_results)

context = ssl.create_default_context()
smtpObj = smtplib.SMTP(login.smtp_server,login.smtp_port)
smtpObj.ehlo()
smtpObj.starttls()
message = ("""Subject: Alert for Community Options Inc -All Mx's - Uplink Packet Loss & Latency\n
The following sites have experienced packet loss above 4 percent during the last 7 days.\n\n
""" + str(final_results))
subject = "Weekly Packet Loss Updates"

smtpObj.login(login.lab_email,login.lab_email_password)
smtpObj.sendmail(login.lab_email,login.company_email, message)

smtpObj.quit()


#interate over dict values
#https://stackoverflow.com/questions/27733685/iterating-over-dict-values

# ignore key error link
#https://stackoverflow.com/questions/15653966/ignore-keyerror-and-continue-program