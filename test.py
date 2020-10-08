def get_averages(file_name):
    global email_body_df # needed to use varibale outside of function.
 
    sites = pd.read_excel(w, sheet_name=None, dtype={'lossPercent':float})

    results = []
    average_latencyMs = []
    average_lossPercent = []
    for office in sites:
        if office not in results:
            results.append(office)
            lossPercent = [sites,office,'lossPercent']
            average_lossPercent.append(lossPercent)
            print(office)
 
    s1 = pd.Series(results, dtype='string', name='Sites')
    s2 = pd.Series(average_latencyMs, dtype='float64', name='latencyMs')
    s3 = pd.Series(average_lossPercent, dtype='float64', name='lossPercent')        

    d = pd.concat([s1,s2,s3],axis=1)
    email_body_df = pd.DataFrame(d)

def send_to_excel(df):
    global writer

    df = email_body_df
    writer = pd.ExcelWriter('averages-'+str(today)+'.xlsx')
    df.to_excel(writer, sheet_name='sites',index=False,float_format="%.2f")
    writer.save()

get_averages(w)

send_to_excel(email_body_df)