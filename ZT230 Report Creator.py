import pandas as pd
import numpy as np
import jinja2

#Creates dataframe from excel document and reads data
location = input('Place your ZT230 Overdue Report in the \\Desktop\\Overdue report\\ZT230\\ folder. Then enter the file name here: ')
data = pd.read_excel("C:\\Users\\jv31651\\Desktop\\Overdue report\\ZT230\\" + location + '.xlsx')
df = pd.DataFrame(data, columns= ['Reported Site', 'Reported Serial ','Shipped Serial ', 'Reported Model','Reported IM', 'Shipped Date', 'Shipped Tracking ', 'Shipped Return Tracking'])


#Separates division and store location from sites then places them into Dataframe
sitelist = []
divisionlist = []
for site in df['Reported Site']:
    divisionlist.append(str(site)[0:2])
    sitelist.append('00' + str(site)[2::])
df['Division'] = np.array(divisionlist)
df['Store'] = np.array(sitelist)

#Creates days over due column
df.loc[df['Shipped Date'] <= (pd.Timestamp.today() + pd.offsets.Day(-14)), 'Days Overdue'] = '2+ Weeks'
df.loc[(df['Shipped Date'] <= (pd.Timestamp.today() + pd.offsets.Day(-7))) & (df['Shipped Date'] > (pd.Timestamp.today() + pd.offsets.Day(-14))),  'Days Overdue'] = '1+ Week'



#Formats column headers and drops Reported Site since it is no longer needed
df = df.rename(columns={'Reported Serial ':'Serial #','Shipped Serial ':'Replaced Serial #', 'Reported Model': 'Model', 'Shipped Tracking ':'Tracking #', 'Shipped Return Tracking':'Return Tracking #', 'Reported IM':'Infra #'})
df = df.drop('Reported Site', 1)


#Formats data in columns
df['Shipped Date'] = df["Shipped Date"].dt.strftime("%m/%d/%y")
df['Tracking #'] = df['Tracking #'].astype(str)
df = df.reindex(columns=['Division', 'Store', 'Model', 'Serial #', 'Replaced Serial #', 'Infra #', 'Shipped Date', 'Tracking #', 'Return Tracking #', 'Days Overdue'])
df['Model'] = 'ZT230'

#Sets the styling of Data. Date >= 28 Days = Red Font: Yellow Highlight
def custom_style(row):
    '''Function to set style based on current reports.
    If Date >= 28 Days ->                   Font = Red,   Highlight = Yellow
    If Date < 28 Days && Date >= 14 Days -> Font = Black, Highlight = Yellow
    If Date < 14 Days && Date >= 7 Days ->  Font = Black, Highlight = Cornflowerblue'''

    font = 'black'
    color = 'white'
    if pd.to_datetime(row.values[6], format='%m/%d/%y') <= (pd.Timestamp.today() + pd.offsets.Day(-14)) and pd.to_datetime(row.values[6], format='%m/%d/%y') <= (pd.Timestamp.today() + pd.offsets.Day(-28)):
        color = 'yellow'
        font = 'red'
    elif pd.to_datetime(row.values[6], format='%m/%d/%y') <= (pd.Timestamp.today() + pd.offsets.Day(-14)):
        color = 'yellow'
    elif pd.to_datetime(row.values[6], format='%m/%d/%y') <= (pd.Timestamp.today() + pd.offsets.Day(-7)) and pd.to_datetime(row.values[6], format='%m/%d/%y') > (pd.Timestamp.today() + pd.offsets.Day(-14)):
        color = 'cornflowerblue'
    return ['background-color: %s; color : %s' % (color, font)]*len(row.values)

#Print to excel
df.style.apply(custom_style, axis=1).to_excel("C:\\Users\\jv31651\\Desktop\\Overdue report\\ZT230\\ZT230 Overdue Return Report " + pd.Timestamp.today().strftime("%Y%m%d") + '.xlsx', index=False, sheet_name= 'Overdue Returns', freeze_panes= (1,0))

