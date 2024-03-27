import os
import pyodbc
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import sys
# r'DBQ=P:\GageTrack_DB\Preventive maintenance\PF-VT-PM-GTData68.mdb;'
#r'DBQ=D:\02. Test data extractor\PF-VT-PM-GTData68.mdb;'



#conn_str = (
#      r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
#      r'DBQ=X:\88-Technology-Innovation-SEA\50-TTD\160-Database\GageTrakDB\PF-VN-CALIBRATION.mdb;'
#      r"PWD=we092uRc1SLKr;"    
#    )

conn_str = 'DSN=pfvn-vm-gagweb;UID=gagread;PWD=Pfvn123456'


equipment = pyodbc.connect(conn_str)
cursor = equipment.cursor()
vnfaserverid='25'
from datetime import datetime, timedelta
Next10days = datetime.now() + timedelta(days=+10)
Next3days = datetime.now() + timedelta(days=+7)
currentdays = datetime.now()

cursor.execute("""
    select Gage_ID,GM_Owner,Description,Current_Location,Calibration_Frequency,Calibration_Frequency_UOM,Next_Due_Date
    from GAGEMGR65.Gage_Master
    where Status IN ('1')
    and Description not like '%RULER%'
"""  )

rows=cursor.fetchall() #read all rows
cursor.close()
equipment.close()
import pandas as pd
df=pd.DataFrame.from_records(rows, columns=["Equip_id","TPU_owner","Description","Location",'Maintenance_freq','Freq_UOM','Due_date'])
df=df[['Due_date',"Equip_id","Description","TPU_owner","Location",'Maintenance_freq','Freq_UOM']]
df=df.sort_values(by='Due_date',ascending=True).reset_index(drop=True)
df=df[df["TPU_owner"]=='PHOTOBATCH']



df10 = df[(df['Due_date'] <= Next10days) & (df['Due_date'] >= Next3days) ].reset_index(drop = True)

df3 =  df[(df['Due_date'] <= Next3days) & (df['Due_date'] >= currentdays) ].reset_index(drop = True)


dfoverdue =  df[(df['Due_date'] <= currentdays) ].reset_index(drop = True)


normalmessage = """<p><b><font color="blue">This is auto email sent if having calibration schedule in next 10 days""";
level1message = """<p><b><font color="blue">This is escalation because your equipment calibration have been already overdued from 1 to 3 days""";
level2message = """<p><b><font color="blue">This is escalation because your equipment calibration have been already overdued more than 3 days""";

photobatchmaillist = 'dhoang@vn.pepperl-fuchs.com;vtnguyen@vn.pepperl-fuchs.com';
photobatchmaillistlv1 = 'dhoang@vn.pepperl-fuchs.com;tgluu@vn.pepperl-fuchs.com';
photobatchmaillistlv2 = "dhoang@vn.pepperl-fuchs.com";"tgluu@vn.pepperl-fuchs.com";



            
################################
if (df3.shape[0] != 0) or (dfoverdue.shape[0] != 0):
    df10 = df10.to_html()
    df3 = df3.to_html()
    dfoverdue = dfoverdue.to_html()
    mailbody = "<!DOCTYPE html>" + "<html><body>"
    #mailbody = mailbody + normalmessage + df10
    mailbody = mailbody + level1message + df3
    mailbody = mailbody + level2message + dfoverdue
    mailbody = mailbody + "</body></html>"

    import win32com.client as win32
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'dhoang@vn.pepperl-fuchs.com'
    mail.Subject = 'Calibration next 10 days'
    mail.HTMLBody = mailbody
    mail.Send()

#conn_str = (
#      r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
#      r'DBQ=X:\88-Technology-Innovation-SEA\50-TTD\160-Database\GageTrakDB\PF-VN-CALIBRATION.mdb;'
#      r"PWD=we092uRc1SLKr;"    
#    )
equipment = pyodbc.connect(conn_str)
cursor = equipment.cursor()
vnfaserverid='25'
from datetime import datetime, timedelta
Next10days = datetime.now() + timedelta(days=+10)
Next3days = datetime.now() + timedelta(days=+7)
currentdays = datetime.now()

cursor.execute("""
    select Gage_ID,GM_Owner,Description,Current_Location,Calibration_Frequency,Calibration_Frequency_UOM,Next_Due_Date
    from GAGEMGR65.Gage_Master
    where Status IN ('1','4')
    and Description like '%RULER%'
"""  )

rows=cursor.fetchall() #read all rows
cursor.close()
equipment.close()
import pandas as pd
df=pd.DataFrame.from_records(rows, columns=["Equip_id","TPU_owner","Description","Location",'Maintenance_freq','Freq_UOM','Due_date'])
df=df[['Due_date',"Equip_id","Description","TPU_owner","Location",'Maintenance_freq','Freq_UOM']]
df=df.sort_values(by='Due_date',ascending=True).reset_index(drop=True)
df=df[df["TPU_owner"]=='PHOTOBATCH']

#df_ruler=df[df["Description"]=='INOX RULER']

dfoverdue1 =  df[(df['Due_date'] <= currentdays) ].reset_index(drop = True)
level3message = """<p><b><font color="blue">This is escalation because your equipment calibration have been already overdued""";
################################
if (dfoverdue1.shape[0] != 0):
    dfoverdue1 = dfoverdue1.to_html()
    mailbody = "<!DOCTYPE html>" + "<html><body>"
    mailbody = mailbody + level3message + dfoverdue1
    mailbody = mailbody + "</body></html>"

    import win32com.client as win32
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    #mail.To = 'dhoang;tqnguyen@vn.pepperl-fuchs.com;tdoan@vn.pepperl-fuchs.com'
    #mail.To = 'dhoang'
    mail.Subject = 'Calibration ruler overdue'
    mail.HTMLBody = mailbody
    mail.Send()
