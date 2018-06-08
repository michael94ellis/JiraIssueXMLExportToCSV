############################################################
### Created by: Michael Ellis, 6/6/18                    ###
###------------------------------------------------------###
### The intended purpose of this script is to export     ###
###  comments from Jira as well as other fields that the ###
###  Jira OOTB application doesn't export for free       ###
###------------------------------------------------------###
### This Powershell Script needs to be ran in STA mode   ###
###  unless your system is using a PowerShell 3 engine   ###
############################################################

# This function will open a file-picker for the user to select their Jira XML Export
Function Get-JiraXMLFile(){ 
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null;
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog;
    $OpenFileDialog.initialDirectory = Get-Location;
    $OpenFileDialog.filter = "XML files (*.xml)|*.xml";
    $OpenFileDialog.ShowDialog() | Out-Null;
    return Get-Content $OpenFileDialog.filename;
}
# This function will open the file save dialong to allow the user to choose location and 
#  name of the converted XML-to-CSV file
Function Get-SaveFile(){ 
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null;
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog;
    $SaveFileDialog.initialDirectory = Get-Location;
    $SaveFileDialog.filter = "CSV files (*.csv)|*.csv";
    $SaveFileDialog.ShowDialog() | Out-Null;
    $SaveFileDialog.filename;
} 
# This function will override the property truncation of Export-CSV by giving all objects 
#  the same properties with $Null values and sorting them too
Function Union-Object ([String[]]$Property = @()) {				
	$Objects = $Input | ForEach {$_}							
	If (!$Property) {ForEach ($Object in $Objects) {$Property += $Object.PSObject.Properties | Select -Expand Name}}
	$Objects | Select ([String[]]($Property | Sort-Object | Select -Unique))
} Set-Alias Union Union-Object

# Invoke the file-picker function and obtain input file 
[Xml]$inputFile = Get-JiraXMLFile;
# Grab all the items we exported, ignore the header info
if ( $inputFile ) {
    #$XmlComments = Select-Xml "//comment()" -Xml $inputFile;
    #$inputFile.RemoveChild($XmlComments);
    $items = Select-Xml "//rss/channel/item" -Xml $inputFile;
}
# Initialize list for items that will be extracted from XML Input File
$list = @(); 

# Iterate over items and grab important info to be put into CSV format
foreach ( $item in $items ){
    $item = $item.Node;
    # Create a new hash object to store data in
    $issue = @{}; 
    #####################################################
    # Jira Issues ought to always have these properties #
    $issue.Key = $item.key.InnerXML;
    $issue.StatusColor = $item.statusCategory.colorName;
    $issue.Status = $item.status.InnerXML;
    $issue.IssueType = $item.type.InnerXML;
    $issue.Resolution = $item.resolution.InnerXML;
    $issue.Summary = $item.summary;
    $issue.Priority = $item.priority.InnerXML;
    $issue.Link = $item.link;
    $issue.Description = $item.description;
    
    #####################################################
    # These issue properties may or may not be there    #
    if( $item.fixVersion){
        $issue.fixVersion = $item.fixVersion;
    }
    # Check for parent
    if( $item.parent){
        $issue.Parent = $item.parent;
    }
    # Check for subtasks
    if( $item.subtasks){
        $incrementalCounter = 0;
        foreach( $task in $item.subtasks.subtask){
            $issue.("subtask"+$incrementalCounter) = $task.InnerXML;  
        }
    }
    
    # Check for comments 
    if ( $item.comments ) {
        # Record the comments with column name/header 
        #  format as follows: comment0 | comment1| ...
        $incrementalCounter = 0;
        # Loop through all comments on the issue
        foreach ( $comment in $item.comments.comment ) {
            ####################################
            ### Parse Comment Text/HTML here ###
            ####################################
            $issue.("comment"+$incrementalCounter) = $comment.InnerXML;
            $incrementalCounter += 1;
        }
    }
    
    # Check for links 
    if ( $item.issuelinks ) {
        # Record the comments with column name/header 
        #  format as follows: links: link0 | link1 | ...
        $incrementalCounter = 0;
        # Loop through all comments on the issue
        foreach ( $link in $item.issuelinks.issuelinktype ) {
            # Record Inward Links
            if ( $link.inwardlinks.issuelink.issuekey ){
                $link = $link.inwardlinks;
                $linkText = $link.description;
                $issue.("link"+$incrementalCounter) = $linkText + " " + $link.issuelink.issuekey.InnerXML;
            }
            # Record Outward Links
            if ( $link.outwardlinks.issuelink.issuekey ){
                $link = $link.outwardlinks;
                $linkText = $link.description;
                $issue.("link"+$incrementalCounter) = $linkText + " " + $link.issuelink.issuekey.InnerXML;
            }
            $incrementalCounter += 1;
        }
    }
    # Custom Fields Contain: Sprint, Epic Link, Rank, Flagged,
    #  Parent Link, Story Points, WSJF, Team Members
    # Check for custom fields
    if( $item.customfields){
        foreach( $field in $item.customfields.customfield){
            if($field.customfieldname -eq "Sprint"){
                $issue.Sprint = $field.customfieldvalues.customfieldvalue.InnerXML;
            }
            if($field.customfieldname -eq "Epic Link"){
                $issue.EpicLink = $field.customfieldvalues.customfieldvalue;
            }
            if($field.customfieldname -eq "Rank"){
                $issue.Rank = $field.customfieldvalues.customfieldvalue;
            }
            if($field.customfieldname -eq "Flagged"){
                # Flagged can only be Impediment
                $issue.Flagged = "Impediment";
            }
           #Write-OutPut $issue;
        }
    }
    
    # Create a Jira Issue object to be added to the list for CSV export
    $list += New-Object –TypeName PSObject –Prop $issue;
}

# Open File Saving window to choose file name and location for the new   
# Union - Override the Export-CSV property truncation and sort properties
$list | Union | Export-CSV -Path (Get-SaveFile) -NoTypeInformation;
