import pandas as pd
import numpy as np
import pyodbc
from datetime import datetime , timedelta

check_time_1 = datetime.now() + timedelta(days = 7) # 7 days
check_time_2 = datetime.now() + timedelta(days = 14) # 14 days
check_time_3 = datetime.now() + timedelta(days = 45) # 45 days
today = datetime.now().date()

# query the data from access database
cnxn =pyodbc.connect( r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=Y:\88-Technology-Innovation-SEA\50-TTD\160-Database\GageTrakDB\PF-VN-CALIBRATION.mdb; PWD=we092uRc1SLKr;')
cursor=cnxn.cursor()
cursor.execute("""
    select Gage_ID , Gage_SN , Description , Status , Model_No , Storage_Location , Current_Location , GM_Owner , Last_Calibration_Date , Next_Due_Date from Gage_Master
    where Status in ('1' , '4')
""")
rows = cursor.fetchall() # get all data meeting the filter query, stored in rows
cursor.close() # close the connection to database
cnxn.close() # close the connection to database

# convert the data from rows to 2D dataframe tables
df = pd.DataFrame.from_records(rows , columns = ['Gage_ID' , 'Gage_SN' , 'Description' , 'Status' , 'Model_No' , 'Storage_Location' ,'Current_Location' , 'GM_Owner' , 'Last_Calibration_Date' , 'Next_Due_Date'])
sts4 = df[df['Status'] == '4'].reset_index(drop = True) # filter all status 4 and stored in sts4
sts4['Last_Calibration_Date'] = sts4.apply(lambda x: x['Last_Calibration_Date'].date() , axis = 1)
sts4['Next_Due_Date'] = sts4.apply(lambda x: x['Next_Due_Date'].date() , axis = 1)

sts1 = df[df['Status'] == '1'].reset_index(drop = True) # filter all status 1 and stored in sts1
sts1_noduedate = sts1[pd.isnull(sts1['Next_Due_Date'])].reset_index(drop = True) # filter all sts1 without next due date 
sts1_withduedate = sts1[~pd.isnull(sts1['Next_Due_Date'])].reset_index(drop = True) # filter all sts1 with next due date
sts1_withduedate['Last_Calibration_Date'] = sts1_withduedate.apply(lambda x: x['Last_Calibration_Date'].date() , axis = 1)
sts1_withduedate['Next_Due_Date'] = sts1_withduedate.apply(lambda x: x['Next_Due_Date'].date() , axis = 1)

# add columns Trigger to classify the data
sts1_withduedate['Trigger'] = sts1_withduedate.apply(lambda x: 'overdue' if x['Next_Due_Date'] < today else 'in 7 days' if x['Next_Due_Date'] <= check_time_1.date() else 'in 14 days' if x['Next_Due_Date'] <= check_time_2.date() else 'in 45 days' if x['Next_Due_Date'] <= check_time_3.date() else 'Others' , axis = 1 )
# filter overdue equipment
overdue = sts1_withduedate[sts1_withduedate['Trigger'] == 'overdue'].reset_index(drop = True)
# filter will be due in 7 days
in7days = sts1_withduedate[sts1_withduedate['Trigger'] == 'in 7 days'].reset_index(drop = True)
# filter will be due in 14 days
in14days = sts1_withduedate[sts1_withduedate['Trigger'] == 'in 14 days'].reset_index(drop = True)
# filter will be due in 45 days
in45days = sts1_withduedate[sts1_withduedate['Trigger'] == 'in 45 days'].reset_index(drop = True)



sts4 = sts4.style.render().replace('\n' , '').replace('<table' , '<table border="1"') 
noduedate = sts1_noduedate.style.render().replace('\n' , '').replace('<table' , '<table border="1"') 
overdue = overdue.style.render().replace('\n' , '').replace('<table' , '<table border="1"') 
in7days = in7days.style.render().replace('\n' , '').replace('<table' , '<table border="1"') 
in14days = in14days.style.render().replace('\n' , '').replace('<table' , '<table border="1"')
in45days = in45days.style.render().replace('\n' , '').replace('<table' , '<table border="1"') 

mailcontent = """<!DOCTYPE html><html><body>"""
mailcontent = mailcontent + """<p><font color="blue"><b>Equipment in calibration (Sts = 4) </b></p>""" + sts4
mailcontent = mailcontent + """<p><font color="blue"><b>Equipment has no overdue date </b></p>""" + noduedate
mailcontent = mailcontent + """<p><font color="blue"><b>Equipment is overdued </b></p>""" + overdue
mailcontent = mailcontent + """<p><font color="blue"><b>Equipment is due in 7 days </b></p>""" + in7days
mailcontent = mailcontent + """<p><font color="blue"><b>Equipment is due in 14 days </b></p>""" + in14days
mailcontent = mailcontent + """<p><font color="blue"><b>Equipment is due in 45 days </b></p>""" + in45days
mailcontent = mailcontent + """</html></body>"""

import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = "unguyen@vn.pepperl-fuchs.com"
mail.Subject = 'Calibration alert'
mail.HTMLBody = mailcontent
mail.Send()
