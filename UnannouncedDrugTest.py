# -*- coding: utf-8 -*-
"""
Created on %(date)s

@author: %(username)s
"""


import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import logging
import time
import getpass

import win32api
import win32con
import win32evtlog
import win32security
import win32evtlogutil
import pyodbc 
#import json
#from datetime import timedelta 
from datetime import date

import pandas as pd
import base64
import requests
from datetime import datetime, timezone, timedelta

"""
Import your libraries here

"""


ph = win32api.GetCurrentProcess()
th = win32security.OpenProcessToken(ph, win32con.TOKEN_READ)
my_sid = win32security.GetTokenInformation(th, win32security.TokenUser)[0]

applicationName = "Unannounced drug test"
eventID = '20-30'

"""
Information        : 10-20 (e.g. events for success of job transaction)
Error              : 20-30 (e.g. events for file missing or different file and job failure)
Warning            : 30-40 (e.g. events for few file copied but some are missing)

"""

os.chdir(os.path.dirname(os.path.abspath(__file__)))
executionusername = getpass.getuser()

start_time = time.time()

#Create and configure logger
LOG_FORMAT = '%(levelname)s %(asctime)s - %(message)s'
logging.basicConfig(filename = os.path.basename(__file__).replace('.py','') + '.log',level = logging.DEBUG, format = LOG_FORMAT, filemode= 'w')
logger = logging.getLogger()

"""
Logging Templates:

logger.debug('This is a debug message')
logger.info('This is an info message')
logger.warning('This is a warning message')
logger.error('This is an error message')
logger.critical('This is a critical message')

"""

try:
    CurrentWorkingDirectory = os.getcwd()
    # -*- coding: utf-8 -*-

    conn = pyodbc.connect('Driver={SQL Server};'
                          'Server=XYZSQLSP1004;'
                          'Database=EDW;'
                          'Trusted_Connection=yes;')
             
    cursor = conn.cursor()
    
    query =''' 
        confidential SQL Query '''
    
    cursor.execute(query)
    
    headings = ['JobInstanceGuId']
    datasets = []
    
    if not os.path.isfile("ListOfGuIds.txt"):
        with open("ListOfGuIds.txt","w") as fp:
            pass
    
    if not os.path.isfile("ErrorLog.txt"):
        with open("ErrorLog.txt","w") as fp:
            pass
    
    payload = {'api_key':'confidential','username':'DSXYZ.api','account':90,'limit':10000}
    r = requests.get('https://api.confidential.com/api/v1/subscription/?', params=payload)
    data = r.json()
    df = pd.DataFrame(data["objects"])
    df = df[df["status"]=="ACTIVE"]
    df_current_pos = pd.DataFrame(columns = ['Shipname','PolestarID', 'longitude', 'latitude', 'timestamp_csp', 'InsideSingaporeArea']) 
    
    
    UtcTime = datetime.now(timezone.utc)
       
    try: 
        for row in cursor:
            InstallationName = str(row[0])
            ChiefEngineerEmailAddres = str(row[1])
            #Person = str(row[2])
            #ChiefEngineerFirstName = str(row[3])
            JobNextDate = row[2]
            #randomdays = [1,2,3,4,5,6,7]
            #f= random.choice(randomdays)
            #g = date.today()
            #x = e-timedelta(days=f)  
            JobGuId = row[3]
            #HSSEManager = row[6]
            RowNumber=row[4]
            #VesselLocalDate = row[8]
            ListOfGuIdFile = open('ListOfGuIds.txt','r')
            contents = ListOfGuIdFile.read()
            for index, row in df.iterrows():
                df2 = pd.DataFrame(row["last_position"])
                df3 = pd.DataFrame(row["ship"])
                shipname = df3.iloc[0]["ship_name"] 
                long = df2.iloc[0]["longitude"]
                LocalVesseltime = UtcTime+ (long*timedelta(minutes = 4)) 
                VesselLocalDate = LocalVesseltime.strftime("%T")
                print(VesselLocalDate)
                #if (InstallationName.upper() == shipname.upper()):
                if (InstallationName.upper() == shipname.upper()):
                    print(InstallationName,shipname)
                    
                    starttime = '09:00:00'
                    enddtime = '19:00:00'
                    
    
                    if RowNumber <= 2 and JobGuId not in contents and VesselLocalDate >= starttime and VesselLocalDate <= enddtime:
                        fromaddr = "Safety@xyz.com"
                        toaddr = ChiefEngineerEmailAddres
                        bcc="aib@xyz.com"
                        # instance of MIMEMultipart 
                        msg = MIMEMultipart() 
                
                # storing the senders email address   
                        msg['From'] = fromaddr 
                
                # storing the receivers email address  
                        msg['To'] = toaddr 
                #msg['Cc'] = cc 
                        msg['Bcc'] = bcc
                
                # storing the subject  
                
                        subject = InstallationName + ' - Engineer\'s unannounced drug test'
                        msg['Subject'] = subject
                        print(subject)
                
                
                        html = '\
            <html> \
              <body> \
               <p>Good Day Engineer,  <br><br>\
             <br><br>\
        Kind Regards,<br>\
        Manager,<br>\
        XYZ Shipping<br>\
        </p> \
              </body> \
            </html>'
        
                
                # string to store the body of the mail 
                #body = "Body_of_the_mail"
                
                # attach the body with the msg instance 
                        msg.attach(MIMEText(html, 'html')) 
                
                
                # creates SMTP session 
                #s = smtplib.SMTP('smtp.office365.com', 587) 
                        s = smtplib.SMTP('smtp.xyz.global',25) 
                # start TLS for security 
                #s.starttls() 
                
                # Authentication 
                
                
                # Converts the Multipart msg into a string 
                        text = msg.as_string() 
                
                # sending the mail 
                
                        receiver_email = toaddr.split(',') +  bcc.split(',')
                        s.sendmail(fromaddr, receiver_email, text) 
                        print("Mail sent")
                #COMMENT ABOVE LINE TO AVOID UNNECESSARY MAILS SENT ON VESSELS
                
                # terminating the session 
                        s.quit() 
                        dataset = dict(zip(headings, JobGuId))
                        datasets.append(dataset)
                
                        today = date.today()
                        formattedtoday = today.strftime("%d/%m/%Y")
                        FileData = JobGuId,formattedtoday
                
                        ListOfGuIds = open('ListOfGuIds.txt','a')
                        print(FileData,file = ListOfGuIds)
                        ListOfGuIds.close()
                        print(ChiefEngineerEmailAddres,InstallationName)
                
                
                
    except Exception as e:
        ErrorLog = open('ErrorLog.txt','a')
        print(e,file = ErrorLog)
        ErrorLog.close()



			

    #else:
     #print(SqlJobInstanceGuId)
               #print(jobInstanceGUID)
									  

    #print(1/0)
except Exception as e:
    logger.exception('ERROR')
    
    msg = MIMEMultipart()
    CurrentWorkingDirectory = os.getcwd()
	
    fromaddr = 'PythonJob@xyz.com'
    
    """Change below to GroupTech@XYZ.com"""
    
    toaddr = 'groupbi@XYZ.com' 
    
    #toaddr = 'GroupBI@XYZ.com' 
    
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = 'Python Script Failed : ' + CurrentWorkingDirectory + ' : ' +  os.path.basename(__file__)
    html = '<html><head><body> <b>Error : </b>' + str(e) \
        + '<br><br> Please check the log file available at : ' \
        + os.path.join(CurrentWorkingDirectory,os.path.basename(__file__).replace('.py','') \
        + '.log') \
        + '<br><br> Program executed by User : ' + executionusername \
        + '<br><br> Program executed from Computer : ' + os.getenv('COMPUTERNAME') \
        + '</body></head></html>'
    
    msg.attach(MIMEText(html, 'html'))

    s = smtplib.SMTP('smtp.xyz.global',25) 
    
    text = msg.as_string() 
    
    receiver_email = toaddr.split(',')
    
    s.sendmail(fromaddr, receiver_email, text)
    
    category = 5
    myType = win32evtlog.EVENTLOG_ERROR_TYPE
    descr = [str(e)]
    data = "Application\0Data".encode("ascii")
     
    win32evtlogutil.ReportEvent(applicationName, eventID, eventCategory=category, 
    eventType=myType, strings=descr, data=data, sid=my_sid)

    
    s.quit() 
else:
    """
    This block will execute once code in try block is completed successfully
    """
finally:
    """
    This block will always execute
    """
    stop_time = time.time()
    dt = stop_time - start_time
    logger.info('Time required to run the code is {time} in Seconds.'.format(time = dt))
    logging.shutdown()
