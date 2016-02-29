# IT_Asset_Management_Tools

Computer Information Collection Utility
Edited: 2/2016
Dan Budris D.C.Budris@gmail.com

CONTEXT
-------
IT asset management tools are frequently geared toward the enterprise, rather than medium sized business.  To bridge this gap I developed tools for automatically inventorying computer hardware and stats and preparing that data for mass upload to our Salesforce IT asset database.  

DESCRIPTION
-----------

CICU is a tool to automate collection of hardware and configuration information from Windows computers in an active directory domain, and provide a graphical interface to view and report on that data.

CICU leverages Windows Managment Instrumentation (WMI) and Group Policy to enable computers to regularly self-report their configuration to a central location.  This information, which is outside of the scope of a regularly configured active-directory computer object, can then be viewed, reported on, and used to gain insight into your assets.  

This project was originally concieved as an intermediate step to create collect data which was then uploaded into salesforce for a more robust reporting system.  

CICU is primarily written in Powershell, some C#, and with use of Windows Forms and some Visual Basic.

HOW IT WORKS
-----------
A powershell logon script is distributed with group policy.

Each computer that runs the logon script will run several WMI queries (comp name, mobo, drives, processer, antivirus, BIOS) and save them in a CSV in the directory specified (in the logon script).

Then the GUI runs (ITAM_GUI) it collects all of the information from the CSV files saved in the directory, allowing you to view the information for individual comptuers as needed.

Additionally, this tool allows you to run Active Directory quieres against the comptuers, returning detailed active directory information without using powershell get-ad statements or filters. 


WHAT YOU NEED
-------------
This tool is designed to be used in a Windows Active Directory domain (funcationality at least 2008) with at least powershell V 3.0.  


CONFIGURATION
-------------
Configure a Group Policy object to run the supplied startup script (InfoCollectorStartupScript.ps1) on the machines you wish to report on. (Located in Policies > Windows Settings > Scripts)

When a computer starts, it will run the script, which reports the data back to a central location.  

Run the main program (CompCollectorGui.ps1) as an administrator; you will be prompted to supply a Domain Controller (for querying additional AD info and adding it to reports), a DataPath (where the data is saved by the startup script), and a ReportPath (where reports are to be saved).  These values will be saved in clixml files in the root directory of the script.  You can change them later in the "options" dropdown menu, or view them in the About menu.

When you open the script, it'll populate with the information reported from each computer.  You can click on one and hit "result" to see the reported information.  

In the options menu, you can hit "produce report" to generate a report that contains all the reported information.  It will be saved to the "reportpath" path.  

TROUBLESHOOTING
---------------

List of computers does not load:
Hit "refresh"
Double check "data path" location contains the reports
Ensure connectivity to the location of the datapath, if it is on a network share
Did you leave enough time for all the comptuers in the environment to report back?

Startup Script is not running:
Double check your GPO.  This is outside the scope of this readme.
Check the Script Execution Policy in your domain; if your script execution policy requires it (allsigned, remotesigned, etc) please digitally sign your script.


BUGS
----
Won't collect antivirus info from Windows Server 2008 or 2012
We need to be able to set the value for the "datapath" in the STARTUP SCRIPT
