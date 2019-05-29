import pandas as pd
import random
import datetime

asana_csv = input('What is the exact .csv file name? \n')
today = datetime.datetime.today()
tomorrow = today + datetime.timedelta(1)
df = pd.read_csv(asana_csv)

############################ Get Rid of Completed Tasks ########################
df = df[df['Completed At'].isnull()]


############################# Reformat Date Stamp ##############################
df['Created At'] = pd.to_datetime(df['Created At'])
df['Created At'] = df['Created At'].dt.strftime('%m/%d/%Y')


############################# Reform Column Names ##############################
df.rename(columns={'Created At':'Start Date','Task ID':'Key','Name':'Title','Assignee':'Assigned To','Due Date':'End Date','Notes':'Description'}, inplace=True)


############################ Add Other Columns #################################
df['Status']='Active'
df['Priority'] = 'Normal'
df['Start Date'] = str(datetime.datetime.strftime(today,'%m/%d/%Y'))
df['End Date'] = str(datetime.datetime.strftime(tomorrow,'%m/%d/%Y'))
df['Assigned To'] = "Garrett Scott<garrettscott@mydealerworld.com>"
df['Duration'] = '1 Day'
df['Depends On'] = ''
df['Start Date Constraint'] = ''


#################### Rearrange The Titles for Wrike Import #####################
df = df[['Key','Title','Status','Priority','Assigned To','Start Date','Duration','End Date','Depends On','Start Date Constraint','Description']]


########################### Give New Key to Tasks ##############################
def random_number(i):
    i = random.randint(1,100000)
    return i

df['Key'] = df['Key'].apply(random_number)


########################## Reset Index & Export ################################
df = df.reset_index(drop=True)
df.to_excel('finished_export.xls')
