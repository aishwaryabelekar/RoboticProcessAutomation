# -*- coding: utf-8 -*-
"""
Created on Tue Dec 19 12:42:59 2023

@author: abelekar
"""

# -*- coding: utf-8 -*-
"""
Created on %(date)s
@author: %(username)s
"""

import datetime
from datetime import date
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import base64
import os
import getpass
import pyodbc
import pandas as pd
import numpy as np
import logging
import time
from datetime import date, timedelta

import win32api
import win32con
import win32evtlog
import win32security
import win32evtlogutil

# Import your libraries here

ph = win32api.GetCurrentProcess()
th = win32security.OpenProcessToken(ph, win32con.TOKEN_READ)
my_sid = win32security.GetTokenInformation(th, win32security.TokenUser)[0]

applicationName = "<<Provide Application Name>>"
eventID = '<<Provide Event ID>> Refer Below'

"""
Information        : 10-20 (e.g. events for success of job transaction)
Error              : 20-30 (e.g. events for file missing or different file and job failure)
Warning            : 30-40 (e.g. events for few file copied but some are missing)
"""

os.chdir(os.path.dirname(os.path.abspath(__file__)))
executionusername = getpass.getuser()

start_time = time.time()

# Create and configure logger
LOG_FORMAT = '%(levelname)s %(asctime)s - %(message)s'
logging.basicConfig(filename=os.path.basename(__file__).replace('.py', '') + '.log', level=logging.DEBUG,
                    format=LOG_FORMAT, filemode='w')
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

    # Your code logic here

except Exception as e:
    logger.exception('ERROR')

    msg = MIMEMultipart()
    CurrentWorkingDirectory = os.getcwd()

    fromaddr = 'PythonJob@xyz.com'

    """Change below to GroupBI@xyz.com"""

    toaddr = executionusername + '@xyz.com'

    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = 'Python Script Failed : ' + CurrentWorkingDirectory + ' : ' + os.path.basename(__file__)
    html = '<html><head><body> <b>Error : </b>' + str(e) \
           + '<br><br> Please check the log file available at : ' \
           + os.path.join(CurrentWorkingDirectory, os.path.basename(__file__).replace('.py', '') \
                          + '.log') \
           + '<br><br> Program executed by User : ' + executionusername \
           + '<br><br> Program executed from Computer : ' + os.getenv('COMPUTERNAME') \
           + '</body></head></html>'

    msg.attach(MIMEText(html, 'html'))

    s = smtplib.SMTP('smtp.xyz.global', 25)

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
    This block will execute once code in the try block is completed successfully
    """
finally:
    """
    This block will always execute
    """
    stop_time = time.time()
    dt = stop_time - start_time
    logger.info('Time required to run the code is {time} in Seconds.'.format(time=dt))
    logging.shutdown()
