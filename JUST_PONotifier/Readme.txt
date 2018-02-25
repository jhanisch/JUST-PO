Version:  1.00b

What is this?
-------------
This is a notification program written for Just Services which will notify users when Purchase Orders
are received into the warehouse.  Notifications come in the form of emails to both the buyer of the
purchase order, and the lead technicial for the job the purchase order is for.

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
                       this is the email address of a user to monitor the system
  Database user info:  This is the connection information for a user to log into the ComputerEase
                       database.  Required information is the User ID (Uid) and Password (Pwd)

This application runs via a scheduled task at desired intervals.  After unpacking the application 
and copying to the desired location, a new Scheduled task will run the job at the desired intervals.  
To create a new Scheduled task, follow the instructions here:  
https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc725745(v=ws.11)
or just google 'windows create scheduled task' for your version of Windows.


Highly recommended things to keep in mind
-----------------------------------------
This application connects to the ComputerEase database and updates data.  Given this, it is higly 
recommended that a separate account be created for this application to use which only has update access to 
the following tables:

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
        <!-- debug : will only email the manager contact email address -->
        <!-- live: will only email the po buyer & lead tech contacts -->
        <!-- monitor: will email both the po buyer & lead tech contacts and the manager contact -->
        <add key="Mode" value="monitor"/>

        <!-- Monitor Email Address  -->
        <!-- required when the Mode is 'monitor'-->
        <add key="MonitorEmailAddress" value=""/>

		<!-- Database connection -->
        <add key="Uid" value=""/>
        <add key="Pwd" value=""/>
        
    </appSettings>
	<!-- Unused at this point -->
    <connectionStrings>
      <add name="JUSTodbc" connectionString="DRIVER={ComputerEase};DBQ=\\JUSTSERVICEINC.local\JS-ACCT\ComputerEase\data\0;UID=MARKH;;NullStrings=0;SERVER=NotTheServer"/>
    </connectionStrings>
</configuration>


History
1.00a - Initial proof of concept.  Identifies PO's to notify and sends simple one line emails
1.00b - Updates email format to include  Job # and Description, Customer, Vendor, Buyer and PO Detail line items.