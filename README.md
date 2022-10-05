# LDAP-QueryAnalyzer

The DC 1644-events can be used to monitor LDAP-traffic and are mostly used to find "bad" queries.

The use-case for this version of the script is for us to be OK with a decomission of some Domain Controllers.
The script will gather all queries to the DC and export them to a CSV-file.

You can that use that CSV to analyze if you might have some system somewhere that have a specifc DC pointed out as an source for LDAP-queries.

### Prerequisites
On the domain controller(s) you need to add a few registry entries (no reboot needed).

By setting "15 Field Engineering" to 5 and "Expensive Search Results Threshold" to 1, all queries are treated as expensive an then logged to Directory Services EventLog with EventID 1644

```
Windows Registry Editor Version 5.00

[HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\NTDS\Diagnostics]
"15 Field Engineering"=dword:00000005

[HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\NTDS\Parameters]
"Expensive Search Results Threshold"=dword:00000001
```
Now you should see 1466-events comming in to the eventlog Directory Services on the Domain Contoller.

You might want to add some extra MB's to the size of the log since this is verbos logging and generates a lot of data.

### Workflow
- Copy the EVTX-files you want to process to a directory (on a powerfull machine).
- Specify that directory as the parameter **$EventLogPath** to the script.
- Run the script **01-CreateFilteredCSV.ps1** (might be a good idea to test with different PowerShell versions and in ISE to get the best performance)
- Now you will end up with CSV-files for each EVTX-file and a summary-file named ***_Filtered-events.csv***
- The summary-file is easy to import into Excel, and in Excel you can create a Pivot Table to analyze what sources and queries you have towards the domain controllers.

### Filtering
Since the CSV-files will become large you can do some basic filtering within the Powershell Script, the parameter ***$filterClientIP*** allows you to filter out specific or wildcarded Client IP's.

The parameter ***$filterLdapQuery*** will filter wildcarded LDAP-queries

The filter(s) will precent events from ending up in the CSV-files.

***Tip:*** Add the IP's of your other domain controllers to prevent internal DC-traffic to end up in the CSV-files

### Analyze LDAP Query Preformance
If you want to use this script to analyze LDAP query performance you should look in to the values of "Search Time Threshold" and "Inefficient Search Results Threshold"

More info can be found in the article [Use Event1644Reader.ps1 to analyze LDAP query performance in Windows Server](https://learn.microsoft.com/en-us/troubleshoot/windows-server/identity/event1644reader-analyze-ldap-query-performance)

Part 2 of the script ***02-CreateExcel.ps1*** can be used to create a spreadsheet with a few tabs for query analysis.

***Important*** If you want to use Part 2 you need to set ***$gatherExtraEventData*** in the first script to ***$true*** before you run it (will take some extra time to run the script)

When you have the new ***_Filtered-events.csv*** with extra event data you can run ***02-CreateExcel.ps1*** to create the file ***_1644Analysis.xlsx***


***FYI:*** The script ***02-CreateExcel.ps1*** is not as refined as part 1 so bugs and problems are mostly from the original script.
