"""
Created on Tue Jul 20 11:02:38 2021

@author: aib
"""

import datetime
from datetime import date
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
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

"""
Import your libraries here

"""

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

    conn = pyodbc.connect('Driver={SQL Server};'
                          'Server=XYZSQLSP1004;'
                          'Database=EDW;'
                          'Trusted_Connection=yes;')
    cursor = conn.cursor()

    query = '''

   WITH WeatherData AS (
    SELECT * 
    FROM EDW.MG.MeteoGroupWeather(NOLOCK) 
    WHERE Latitude NOT LIKE '%E%' AND Longitude NOT LIKE '%E%'
    AND IsCurrent = 1 AND IsDeleted = 0
)
, VoyageWeatherInfo AS (
    SELECT
        iv.VoyageNumber AS [Voyage Number (IMOS)],
        CAST(ai.ShipName AS NVARCHAR(MAX)) AS [Ship Name],
        CAST(iv.OperatorCoordinator AS NVARCHAR(MAX)) AS [Operator Initials],
        CAST(ai.DepartureLocation AS NVARCHAR(MAX)) AS [Departure Location (MeteoGroup)],
        CAST(ai.DestinationLocation AS NVARCHAR(MAX)) AS [Destination Location (MeteoGroup)],
        iv.StartDateTime AS [Voyage Start],
        iv.EndDateTime AS [Voyage End],
        CAST(wd.ValidAt AS NVARCHAR(50)) AS [Time of Weather],
        WaveAlert = CASE WHEN CAST(wd.TotalWaveHeightInMeters AS DECIMAL(3,1)) >= 5.0 THEN 'Yes' ELSE 'No' END,
        CAST(wd.TotalWaveHeightInMeters AS DECIMAL(3,1)) AS [Wave Height (Meters)],
        iv.DataQuality AS [Data Quality],
        WeatherGuid = wd.Guid,
        iv.Guid
    FROM WeatherData wd
    JOIN EDW.IMOS.IMOSVoyages iv
        ON wd.Latitude = iv.Latitude
        AND wd.Longitude = iv.Longitude
        AND wd.ValidAt BETWEEN iv.StartDateTime AND iv.EndDateTime
        AND iv.DataQuality = 'Voyage is OK'
    LEFT JOIN EDW.MG.AvailabilityInfo ai
        ON iv.Guid = ai.IMOSVoyageGuid
        AND wd.Latitude = ai.Latitude
        AND wd.Longitude = ai.Longitude
        AND wd.ValidAt BETWEEN ai.DepartureTime AND ai.ArrivalTime
    WHERE wd.ValidAt BETWEEN '2022-11-01 00:00:00.000' AND '2022-11-30 23:59:59.999'
)
, DetailedVoyageInfo AS (
    SELECT 
        vwi.*,
        ISNULL(CONVERT(VARCHAR(50), portCalls.DepartureTime, 20), CAST(CAST(vwi.StartDateTime AS DATE) AS NVARCHAR(50)) + ' 00:00:00.000') AS VoyageStart,
        ISNULL(CONVERT(VARCHAR(50), portCalls.ArrivalTime, 20), CAST(CAST(vwi.EndDateTime AS DATE) AS NVARCHAR(50)) + ' 23:59:59.999') AS VoyageEnd,
        CONVERT(DATE, wd.ValidAt) AS WeatherDay
    FROM VoyageWeatherInfo vwi
    LEFT JOIN EDW.EDWDataCube.Staging2.Fact_Stage_D_Port_Calls portCalls
        ON vwi.Guid = portCalls.MeteoGroupShipID
)
SELECT *
FROM DetailedVoyageInfo
WHERE WeatherDay BETWEEN VoyageStart AND VoyageEnd;


    '''
    df = pd.read_sql_query(query, conn)
    conn.close()
    today = date.today()

    df['Voyage Start'] = pd.to_datetime(df['Voyage Start'])
    df['Voyage End'] = pd.to_datetime(df['Voyage End'])
    df['Time of weather'] = pd.to_datetime(df['Time of weather'])

    df['WindAlert'] = np.where(df['Wind (knots)'] >= 20, 'Yes', 'No')

    df = df[['Voyage number (IMOS)', 'Ship name', 'Operator initials', 'Departure (MeteoGroup)',
             'Destination (MeteoGroup)', 'Voyage Start', 'Voyage End', 'Time of weather',
             'Wind (knots)', 'WindAlert', 'WaveAlert', 'Wave height (meters)', 'Data quality']]

    df['Voyage Start'] = df['Voyage Start'].dt.strftime('%Y-%m-%d %H:%M:%S')
    df['Voyage End'] = df['Voyage End'].dt.strftime('%Y-%m-%d %H:%M:%S')
    df['Time of weather'] = df['Time of weather'].dt.strftime('%Y-%m-%d %H:%M:%S')

    df.to_csv('MG_Weather_Alerts_' + str(today) + '.csv', index=False)
    logger.info('Data exported successfully to CSV file')

except Exception as e:
    logger.error('Error occurred: ' + str(e))

finally:
    end_time = time.time()
    execution_time = end_time - start_time
    logger.info('Total execution time: ' + str(execution_time) + ' seconds')
    logger.info('--------------------------------------------------------')

# Send log file as an attachment in an email
try:
    sender_email_address = "<<Provide Sender Email Address>>"
    receiver_email_address = "<<Provide Receiver Email Address>>"
    subject = "Weather Alert Script Execution - XYZ"

    msg = MIMEMultipart()
    msg['From'] = sender_email_address
    msg['To'] = receiver_email_address
    msg['Subject'] = subject

    body = "Weather Alert Script Execution - XYZ\n\n"
    body += "Script Name: " + os.path.basename(__file__) + "\n"
    body += "Execution Time: " + str(execution_time) + " seconds\n\n"
    body += "Log File Attached."

    msg.attach(MIMEText(body, 'plain'))

    attachment = open(os.path.basename(__file__).replace('.py', '') + '.log', 'rb')
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= " + os.path.basename(__file__).replace('.py', '') + '.log')

    msg.attach(part)

    server = smtplib.SMTP('<<Provide SMTP Server>>', 587)
    server.starttls()
    server.login(sender_email_address, "<<Provide Sender Email Password>>")
    text = msg.as_string()
    server.sendmail(sender_email_address, receiver_email_address, text)
    server.quit()

except Exception as e:
    logger.error('Error sending email: ' + str(e))
