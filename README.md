These Python scripts are programs designed for initiating and managing unannounced drug tests for chief engineers on vessels, likely in a maritime or shipping context. Below is a breakdown of the key components and functionalities:

Logging and Error Handling:

The script begins with the setup of logging functionality to record different log levels (debug, info, warning, error, critical) in a log file.
Database Connection:

Connects to a SQL Server database using pyodbc to execute a confidential SQL query. The nature of the query is not provided.
Data Processing:

Processes the data obtained from the SQL query, including details about installations, chief engineers, job dates, job GUIDs, row numbers, etc.
Vessel Position Data:

Retrieves vessel position data from an external API (Polestar Global API) and extracts relevant information such as ship names, longitudes, and timestamps.
Email Notification:

Sends email notifications to chief engineers under specific conditions, such as when certain criteria related to the vessel position, job details, and previous tests are met.
The email content includes instructions for the chief engineer to conduct an unannounced drug test and provide results within a specified time frame.
File Handling:

Manages two files, "ListOfGuIds.txt" and "ErrorLog.txt," for recording processed job GUIDs and any encountered errors, respectively.
Error Notification:

In the event of an exception or error during script execution, an error notification email is sent to a specified email address. Additionally, an entry is made in the Windows Event Log.
Execution Time Logging:

Logs the execution time of the script, indicating the time taken to run the entire code.
