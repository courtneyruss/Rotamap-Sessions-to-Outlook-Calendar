import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime
from datetime import timedelta
from datetime import tzinfo
from dateutil import tz
import pytz
import requests
from dateutil.relativedelta import relativedelta, FR
from urllib.parse import urljoin
import os
import win32com.client
import time
import sys
import numpy as np
import re



with open(r'calendar_sessions_info.txt', 'r') as f:
    lines = f.readlines()
clinician = lines[0].strip()
print(clinician)
print(type(clinician))
clin_email = lines[1].strip()
new_dir = lines[2].strip()
site_url = lines[3].strip()
login_url = lines[4].strip()
prox = lines[5].strip()
proxies={"https":prox}
tokenfile = lines[6].strip()
xmlfile = lines[7].strip()

### Part 1: Delete existing appointments

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
recipient = namespace.createRecipient(clin_email)
resolved = recipient.Resolve()
calendar = namespace.GetSharedDefaultFolder(recipient, 9)

today = datetime.today()
today = today.strftime('%m/%d/%Y %H:%M %p')

# Filter out appointments that occur before today
appointments = calendar.Items.Restrict("[Start] >= today'")

global countvar
countvar = 1

while countvar > 0:
    countvar = 0
    for appointment in appointments:
        print(appointment.Subject)
        if (appointment.Categories == "Green Category" or appointment.Categories == "Orange Category"):
            appointment.Delete()
            print("Deleted")
            countvar = countvar + 1
        else:
            print("Not Deleted")
    print(countvar)



    


### Part 2: Get data from CLW Rota
#Get username and password
password = os.environ.get('clwrotamasterpw')
username = os.getenv('clwrotamaster')


##Get token XML
resp = requests.post(login_url, data ={'username':username, 'password': password}, proxies=proxies, allow_redirects=True)
with open(tokenfile, 'wb') as file:
    file.write(resp.content)

#pull token from XML
tree1=ET.parse(tokenfile)
root1=tree1.getroot()

for token in root1.iter('token'):
    global tok
    tok=(token.text)
print(tok)
##generate URL for calendar info
today=datetime.today()
date1= today.strftime("%Y-%m-%d") + "&"
dt2= today + timedelta (weeks = 8)
date2 = dt2.strftime("%Y-%m-%d")

finalurl=site_url + tok + '/' + 'person_rota/?from_date=' + str(date1) + 'to_date=' + str(date2)

print(finalurl)

##get response

response=requests.get(finalurl, proxies=proxies, allow_redirects=True)
with open(xmlfile, 'wb') as file:
    file.write(response.content)


#Parse XML
tree=ET.parse(xmlfile)
root=tree.getroot()
dict=[]

#Define variables
for item in root.iter('person_rota_item'):
    name=item.find('./person').text
    oncall = item.find('./location').text
    loctype = item.find('./location_type').text
    start=str(item.find('./date').text) + " " + str(item.find('./start_time').text)
    startdt = datetime.strptime(start, "%Y-%m-%d %H:%M:%S")
    session=str(item.find('./session').text)
    speciality = item.find('./speciality')
    title = item.find('./title')
    if (speciality is None) and (title is None):
        speciality = oncall
    elif (speciality is None) and (title is not None):
        speciality = title.text
    else: speciality = speciality.text
    if session == 'am':
        duration = 240
    else: duration = 300
    timesession= start + session + speciality
    #Put variables into dictionary
    dict.append([name, oncall, loctype, startdt, session, timesession, duration, speciality])


### Part 3: Generate Dataframe

## Dictionary to dataframe
df=pd.DataFrame(dict, columns = ['Name', 'Location', 'Location_Type', 'Start', 'Session', 'TimeSession', "Duration", "Speciality"])
df['Start']=pd.to_datetime(df['Start'])
df['StartDate']=df['Start'].dt.date
df['Day'] = df['Start'].dt.strftime('%A')

merged_df = df.groupby("TimeSession").agg({"Name": lambda x: ", ".join(x),
                                        "Start": "first",
                                        "Location": "first",
                                        "Speciality": "first",
                                        "Session": "first",
                                        "Location_Type": "first",
                                        "Duration" : "first",}).reset_index()



### Part 4: Add clinical sessions

#Filter df for clinician
clin_sess_df = merged_df.query("Name.str.contains(@clinician) and Location_Type=='Standard' and Location != 'NCD'")

clin_sess_df = clin_sess_df.replace(clinician, ' ', regex=True)

clin_sess_df = clin_sess_df.replace(',', '', regex=True)

clin_sess_df.loc[:, 'Name'] = clin_sess_df['Name'].str.lstrip()

pattern = r'\(Fellow\)'
filtered_df3.loc[:, 'Name'] = filtered_df3['Name'].str.replace(pattern, '', regex=True)
filtered_df3.to_csv('data4.csv') ##all clinical sessions

#Add to calendar
for index, row in filtered_df3.iterrows():
    print(row['Location'])
    st = (str(row['Start']))
    stdt = datetime.strptime(st, "%Y-%m-%d %H:%M:%S")
    dt_utc = stdt.replace(tzinfo=pytz.UTC)
    local_zone = tz.tzlocal()
    dt_local = dt_utc.astimezone(local_zone)
    appt = calendar.Items.Add(1)
    appt.Start = dt_local
    appt.Location = str(row['Location'])
    appt.Subject = str(row['Speciality'])
    appt.Duration = row['Duration']
    appt.Categories = 'Orange Category'
    if row['Name']=='':
        appt.Body = 'From CLW Rota'
    else:
        appt.Body = "With " + row['Name'] + "\nFrom CLW Rota"
    appt.ReminderSet = False
    appt.Save()


### Part 5: Add On Call
## Filter for on call
filtered_df=df.query('Location in ["Registrar On call", "Consultant On Call"] and Session in ["am", "eve"]')
pattern = r'\(Fellow\)'
filtered_df_copy = filtered_df.copy()
filtered_df_copy.loc[:, 'Name'] = filtered_df_copy['Name'].str.replace(pattern, '', regex=True)
pattern2 = r'\(Pain\)'
filtered_df_copy2 = filtered_df_copy.copy()
filtered_df_copy2.loc[:, 'Name'] = filtered_df_copy2['Name'].str.replace(pattern2, '', regex=True)
filtered_df_copy2.loc[:, 'Oncall2'] = filtered_df_copy2['Location'].str.split().str.get(0)
filtered_df_copy2.loc[:, 'Name'] = filtered_df_copy2['Oncall2'] + ": " + filtered_df_copy2['Name']
filtered_df_copy2['Title'] = filtered_df_copy2['Location'] + " " + filtered_df_copy2['Session']
filtered_df_copy3 = filtered_df_copy2.copy()
filtered_df_copy3.loc[:, 'Title'] = filtered_df_copy3['Title'].str.replace('am', '- Day', regex=True)
filtered_df_copy3.loc[:, 'Title'] = filtered_df_copy3['Title'].str.replace('eve', ' - Night', regex=True)


merged_df = filtered_df_copy3.groupby("TimeSession").agg({"Name": lambda x: ", ".join(sorted(x)),  # Sort names alphabetically
    "Start": "first",
    "Day": "first",
    "Location": lambda x: "".join(x),
    "Title": "first",
    "Location_Type" : "first"}).reset_index()


filtered_df2 = merged_df.query('Name.str.contains(@clinician)')
filtered_df2.to_csv('data3.csv')
filtered_2 = filtered_df2.copy()

print(filtered_2)
filtered_2.to_csv('data2.csv')

replace1 = str("Consultant: " + clinician + ",")
print(replace1)
replace2 = str(", Consultant: " + clinician)
print(replace2)
replace3 = str("Consultant: " + clinician)
print(replace3)
filtered_2.loc[:, 'Name'] = filtered_2['Name'].replace(replace1, '', regex=True)
filtered_2.loc[:, 'Name'] = filtered_2['Name'].str.replace(replace2, '', regex=True)
filtered_2.loc[:, 'Name'] = filtered_2['Name'].str.replace(replace3, '', regex=True)
filtered_2.loc[:, 'Name'] = filtered_2['Name'].str.replace('Consultant', 'Fellow', regex=True)
filtered_2.loc[:, 'Name'] = filtered_2['Name'].str.replace(' , ', ', ', regex=True)

print(filtered_2)
filtered_2.to_csv('data.csv')


## Add call to calendar
for index, row in filtered_2.iterrows():
    st = (str(row['Start']))
    stdt = datetime.strptime(st, "%Y-%m-%d %H:%M:%S")
    dt_utc = stdt.replace(tzinfo=pytz.UTC)
    local_zone = tz.tzlocal()
    dt_local = dt_utc.astimezone(local_zone)
    appt = calendar.Items.Add(1)
    appt.Start = dt_local
    appt.AllDayEvent = True
    appt.ReminderSet = False
    appt.Categories = 'Green Category'
    appt.Body = "From CLW Rota"
    if (row['Name'] is None):
        appt.Location = ""
        appt.Subject = "Consultant on Call"
    elif ',' in (row['Name']):
        appt.Location = str(row['Name'])
        appt.Subject = "Consultant on Call (Day)"
    elif (row['Day'] in ['Saturday', 'Sunday']):
          appt.Location = str(row['Name'])
          appt.Subject = "Consultant on Call (Night)"
    else:
        appt.Location = str(row['Name'])
        appt.Subject = "Consultant on Call"
    print("Added " + str(appt.Subject) + " " + str(appt.Start))
    appt.Save()
       
### Part 6: Add Liver Call
## Filter for liver call
liver_df = df.query('Location == "Liver On Call" and Name == @clinician and Session == "eve"')

## Add liver call to calendar
for index, row in liver_df.iterrows():
    st = (str(row['Start']))
    stdt = datetime.strptime(st, "%Y-%m-%d %H:%M:%S")
    dt_utc = stdt.replace(tzinfo=pytz.UTC)
    local_zone = tz.tzlocal()
    dt_local = dt_utc.astimezone(local_zone)
    appt = calendar.Items.Add(1)
    appt.Start = dt_local
    appt.AllDayEvent = True
    appt.ReminderSet = False
    appt.Subject = "Liver on Call"
    appt.Body = "From CLW Rota"
    appt.Categories = "Green Category"
    appt.Save()


### Part 7: Add nonclinical and admin
## Filter for nonclinical and admin
nonclindf = df.query('Location in ["NCD", "Admin"] and Name == @clinician')

## Add non clin and admin to calendar
for index, row in nonclindf.iterrows():
    st = (str(row['Start']))
    stdt = datetime.strptime(st, "%Y-%m-%d %H:%M:%S")
    dt_utc = stdt.replace(tzinfo=pytz.UTC)
    local_zone = tz.tzlocal()
    dt_local = dt_utc.astimezone(local_zone)
    appt = calendar.Items.Add(1)
    appt.Start = dt_local
    appt.AllDayEvent = False
    appt.Duration = 30
    appt.BusyStatus = 0
    appt.ReminderSet = False
    appt.Subject = str(row['Location']) + " " + str(row['Session'])
    appt.Body = "From CLW Rota"
    appt.Categories = "Orange Category"
    appt.Save()



print('Done')
