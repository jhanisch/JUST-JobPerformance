Version:  1.00

What is this?
-------------
This is a notification program written for Just Services which is designed to run on a daily basis
to notify users when new jobs have been created.  Notifications come in the form of emails to primary 
and secondary contacts for the customer the job was created for.  A full summary of all new jobs
will be sent to each email in the list of Executives defined in the configuration file.

Requirements
------------
This application relies on the ODBC system database connection to the ComputerEase database.  Therefore
this application must be installed on the server which hosts the ComputerEase database.  At the time of
this writing this is the 'js-acct' server at Just Services.

How to install this application
---------------------------
After unpacking the contents of this package, copy all the contents of this to a folder where they
can live on the same server as the ComputerEase database (js-acct).  After unpacking and copying to 
a folder, configure the app.config with values for the following required information:
  From Email account:  This is the SMTP information which will send email to the users
  Mode:                Valid values are: debug, live and monitor.  
                       'debug' will only email the manager contact email when po's are received.  
					   'live' will email only the buyer and lead technicians when a purchase order
					   received
					   'monitor' will email the buyer, lead tech and monitor email addresses when
					   a purchase order is received.
  Montior Email Address:  If desired, and the system is configured in either 'debug' or 'monitor' mode
                       this is the email address of a user to monitor the system.  Multiple email addresses
					   can be emailed, separate each email address with a ;
  Database user info:  This is the connection information for a user to log into the ComputerEase
                       database.  Required information is the User ID (Uid) and Password (Pwd)

This application runs via a scheduled task at desired intervals.  After unpacking the application 
and copying to the desired location, a new Scheduled task will run the job daily.  
To create a new Scheduled task, follow the instructions here:  
https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc725745(v=ws.11)
or just google 'windows create scheduled task' for your version of Windows.


Highly recommended things to keep in mind
-----------------------------------------
This application connects to the ComputerEase database and updates data.  Given this, it is higly 
recommended that a separate account be created for this application to use which only has update access to 
read from the following tables:
jcjob
customer
premployee

Update access is only required on:
jcjob

Given that the database user and password are stored in the configuration file, it is highly recommended
to encrypt the app.config file after configuration and testing.  To do this, read:
https://docs.microsoft.com/en-us/dotnet/framework/data/adonet/connection-strings-and-configuration-files



Blank app.config
----------------
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <appSettings>
    <!-- Email  -->
    <add key="FromEmailAddress" value="Notifications@justserviceinc.com"/>
    <add key="FromEmailPassword" value="Rack6451"/>
    <add key="FromEmailSMTP" value="smtp.office365.com"/>
    <add key="FromEmailPort" value="587"/>

    <!-- Mode -->
    <!-- Valid modes are:  -->
    <!-- debug : will only email the Executive email addresses -->
    <!-- live: will only email the customer primary and secondary contacts -->
    <!-- monitor: will email both the customers primary contact, secondary contact (if defined) and ExecutiveEmailAddresses -->
    <add key="Mode" value="debug"/>

    <!-- Monitor Email Address  -->
    <!-- required when the Mode is 'monitor'-->
    <!-- Separate multiple email addresses with a ; or ,-->
    <add key="ExecutiveEmailAddresses" value=""/>

    <add key="Uid" value=""/>
    <add key="Pwd" value=""/>
  </appSettings>
</configuration>


History
-------
1.00 - Initial go-live delivery.
1.0.1.0 - Update with holiday date logic
1.0.1.1 - Added Veterans Day Observed
