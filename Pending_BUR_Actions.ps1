#* FileName: Pending_BUR_Actions.ps1
#*=============================================
#* Script Name: Pending_BUR_Actions.ps1
#* Description: This script will send the consolidated mail on the pending BUR actions assigned.
#* Created: 03-July-2018
#* Author: Vimaleshwara Gajanana
#*=============================================

Add-Type -Path 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll'
Add-Type -Path 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll'
Add-Type -Path 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.WorkflowServices.dll'

#Email settings
$strFrom = "vimaleshwara.gajanana@unisys.com"
$strTo = "vimaleshwara.gajanana@unisys.com"
$strCc = "vimaleshwara.gajanana@unisys.com"
$strSMTPServer ="na-mailrelay-t3.na.uis.unisys.com"
$strSubject = "IT BUR Action Reminder"

$logfile = "Pending_BUR_Actions.log"

function fnSendMail ($message)
{
	Send-MailMessage -To $strTo -From $strFrom -Cc $strCc -Subject $strSubject -BodyasHTML $message -SmtpServer $strSMTPServer
}

function writeToLog ($message)
{
	$temp = (get-date -Format u) + " - " + $message
	Add-Content -Path $logfile -Value $temp
}

#Enter the site URL, List name
$siteURL = "https://unisyscorp.sharepoint.com/sites/global_operations"
$listname = "BUR List of Actions"

Write-Host "Enter the SPO credentials"
$Cred = Get-Credential

$spCtx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL) 
$spCredentials = New-Object  Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.UserName, $Cred.Password)  
$spCtx.Credentials = $spCredentials

$spWeb = $spCtx.Web
$spLists = $spWeb.Lists
$spCtx.Load($spLists)
$spCtx.ExecuteQuery()

$spList= $spLists.GetByTitle($listname)
$spCtx.Load($spList)
$spCtx.ExecuteQuery()

$query = New-Object Microsoft.SharePoint.Client.CamlQuery
$query.ViewXml = "<View><Query><Where><And><Neq><FieldRef Name='Action_x0020_Status'/><Value Type='Choice'>Completed</Value></Neq><Leq><FieldRef Name='Action_x0020_Due_x0020_Date' /><Value Type='DateTime'><Today OffsetDays='+30' /></Value></Leq></And></Where></Query><OrderBy><FieldRef Name='Action_x0020_Due_x0020_Date' Ascending='FALSE' /></OrderBy></View>"

$items = $spList.GetItems($query)
$spCtx.Load($items)
$spCtx.ExecuteQuery()

$Assignees = @()

foreach($item in $items)
{
	#Get the assignee names
	$Assignees += $item.FieldValues["Action_x0020_Assigned_x0020_to"].Email
}

$UniqueAssignees = $Assignees | Select-object -Unique | Sort-Object

#Write-Host $UniqueAssignees

foreach($UniqueAssignee in $UniqueAssignees)
{
	$strBody = "<p>Hi,</p><p>From the past BUR Meetings there are few actions awaiting your response."
	$strTable = "<table style='font-family:Calibri;font-size:11pt' border=1 cellspacing=0 cellpadding=5><tr><th>BUR Name</th><th>Action Description</th><th>Action Due Date</th><th>Action Status</th></tr>"
	foreach ($item in $items)
	{
		if($item.FieldValues["Action_x0020_Assigned_x0020_to"].Email -contains $UniqueAssignee)
		{
			$strTable = $strTable + "<tr><td>$($item.FieldValues[""BUR_x0020_Name""])</td><td>$($item.FieldValues[""Action_x0020_Description""])</td><td>$($item.FieldValues[""Action_x0020_Due_x0020_Date""].GetDateTimeFormats(""d"")[0])</td><td>$($item.FieldValues[""Action_x0020_Status""])</td></tr>"
		}
	}
	$strBody += $strTable + "</table>"
	$strBody += "<p>Please click <a href='https://unisyscorp.sharepoint.com/sites/global_operations/Lists/List%20of%20Action/My%20Open%20Items.aspx'>here</a> to change the status, due date and enter comments.  You will receive regular reminders to update this action item until it is marked as closed.</p>"
	$strBody += "<p>Thank you and Best Regards,<br>Cris Aguiar</p></font>"

	#Write-Host $strBody
	
	fnSendMail $strBody
	Exit(0)
}




Write-Host "Completed"