# JiraIssueXMLExportToCSV
This project provides an easy and free way to export all the fields that Jira Export cannot do in CSV format on its own. 

This will provide you with:
    <ul>
    <li>Comments</li>
    <li>Issue Link Types and Keys</li>
    <li>Status Color/Category</li>
    <li>Rank</li>
    <li>A link to the issue</li>
    </ul>

To use this, export the XML for the Jira issues you want in CSV.

Then download these files:
    RunTheConverter.cmd - It's one line, it runs the powershell script in STA mode and by-passes security
                          You can rename the RunTheConverter.cmd file
    XMLToCSVConverter.ps1 - This is a PowerShell 2.0 script that opens a File-Picker window and parses XML
                          into PSObjectes then opens a File-Saver window and saves them as a CSV file

Run the RunTheConverter.cmd file to use this tool.


This is still a work in progress.
