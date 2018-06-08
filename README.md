# JiraIssueXMLExportToCSV
This project provides an easy and free way to export all the fields that Jira Export cannot do in CSV format on its own. 

To use this, export XML the Jira issues you want.

Then download these files:
    RunTheConverter.cmd - It's one line, it runs the powershell script in STA mode and by-passes security
                          You can rename the RunTheConverter.cmd file
    XMLToCSVConverter.ps1 - This is a PowerShell 2.0 script that opens a File-Picker window and parses XML
                          into PSObjectes then opens a File-Saver window and saves them as a CSV file

Run the RunTheConverter.cmd file to use this tool.


This is still a work in progress.
