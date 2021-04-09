# Windows Updates Report Generator
# Written By: Steve Lunn (s.m.lunn@wolftech.f9.co.uk)
# Downloaded From: https://github.com/Gilgamoth

# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.

# You should have received a copy of the GNU General Public License
# along with this program.  If not, see https://www.gnu.org/licenses/

$App_Version = "2021-04-09-1030"

Clear-Host
Set-PSDebug -strict
#$ErrorActionPreference = "SilentlyContinue"
[GC]::Collect()

# **************************** LOAD CONFIG FILE ****************************

# If it exists, load the config file.
if(Test-Path wsus-report-config.ps1)
{
 . .\wsus-report-config.ps1
}
else
{
 Write-Host "Error! " -NoNewline -ForegroundColor Red
 Write-Host "wsus-report-config.ps1 not found in script location"
 Exit
}

# *************************** FUNCTION SECTION ***************************

# ****************************** CODE START ******************************

# Declare Variables

$StartTime = (get-date).ToString("dd-MM-yyyy HH:mm:ss")
$ComputersNotCheckingIn = @()
$ComputersNotCheckingInCount = 0
$ReportFile = $Cfg_BaseReportFolder + "index.html"
$Emailbody =  $Cfg_BaseReportFolder + "emailbody.txt"
$Today=(get-date).ToString("yyyy-MM-dd HH:mm:ss")
$UnassignedComputers = @()
$ComputersNotPatching = @()
$ExcludedSvrs = ""
$wsus = ""

"<HEAD><TITLE>WSUS Report "+(get-date).ToString("yyyy-MM-dd")+"</TITLE></HEAD><BODY><TABLE border=1 width=1280><TR><TH>Server Name</TH><TH>FQDN</TH><TH>Last Check In Time</TH><TH>Total Patches</TH><TH>Required Patches</TH><TH>Failed Patches</TH><TH>Approved Patches</TH></TR>" | Out-file -FilePath $ReportFile -encoding ASCII

If ($Cfg_ExcludedServerFile) {
	If(Test-Path $Cfg_ExcludedServerFile) { 
		$ExcludedSvrs = Get-Content $Cfg_ExcludedServerFile
	} Else {
		Write-Host "Error! " -NoNewline -ForegroundColor Red
		Write-Host "$Cfg_ExcludedServerFile not found but specified"
		Exit
	}
}

If ($Cfg_UsersFile) {
	If(Test-Path $Cfg_UsersFile) { 
		$UsersFile = Import-Csv -Path  $Cfg_UsersFile
	} Else {
		Write-Host "Error! " -NoNewline -ForegroundColor Red
		Write-Host "$Cfg_UserFile not found but specified"
		Exit
	}
}

if(Test-Path $Cfg_BaseReportFolder) { 
} Else {
	Write-Host "Error! " -NoNewline -ForegroundColor Red
	Write-Host $Cfg_BaseReportFolder " not available"
	Exit
}

# Connect to WSUS Server
[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | out-null 

If($Cfg_WSUSServer) {
	$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($Cfg_WSUSServer, $Cfg_WSUSSSL, $Cfg_WSUSPort)
} else {
	$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer()
}

# Set up Computer Scope

$ComputerScope = new-object Microsoft.UpdateServices.Administration.ComputerTargetScope
$ComputerScope.IncludedInstallationStates = [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::All
$ComputerScope.IncludeDownstreamComputerTargets = $true
$ComputerScope.IncludeSubgroups = $true

if($Cfg_WSUSTarget) {
	$wsus.GetComputerTargetGroups() | where {$_.Name -eq $Cfg_WSUSTarget} | %{$ComputerScope.ComputerTargetGroups.Add($_)} | out-null
} else {
	$wsus.GetComputerTargetGroups() | where {$_.Name -ne "Unassigned Computers"} | %{$ComputerScope.ComputerTargetGroups.Add($_)} | out-null
}

$Computers = $wsus.GetComputerTargets($ComputerScope) | sort -Property FullDomainName

# Set Updates Scope
$UpdateScope = new-object Microsoft.UpdateServices.Administration.UpdateScope
$UpdateScope.IncludedInstallationStates = [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::NotInstalled, [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::Downloaded, [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::Failed

# Exclude Definition Updates Classification
$wsus.GetUpdateClassifications() | where {$_.Title -ne "Definition Updates"} | %{$UpdateScope.Classifications.Add($_)} | out-null

# Include Windows and Forefront Categories only
#$wsus.GetUpdateCategories() | where {$_.Title - contains "Windows","Forefront"} | %{$UpdateScope.Categories.Add($_)} | out-null

# Process for each computer in Computer Scope

ForEach($Computer in $Computers) {
	$WsusComputerName = $Computer.FullDomainName
	$WsusTotalPatchCount = $Computer.GetUpdateInstallationSummary()
	$WsusFailedPatchCount = $WsusTotalPatchCount.FailedCount
	$WsusReqPatchCount = ($Computer.GetUpdateInstallationInfoPerUpdate($UpdateScope)).Count
	$WsusLastCheckin = $Computer.LastReportedStatusTime
	$WsusTotalUpdateCount = $WsusTotalPatchCount.NotInstalledCount + $WsusTotalPatchCount.DownloadedCount
	$WsusReqUpdateCount = $WsusReqPatchCount.NotInstalledCount + $WsusReqPatchCount.DownloadedCount + $WsusTotalPatchCount.FailedCount
	[int]$PatchGroup="0"
	if ($Computer.ComputerTargetGroupIds[$PatchGroup].guid -eq "a0a08746-4dbe-4a37-9adf-9e7652c0b421") {
		$PatchGroup=1
	}
	$WsusPatchGroup = ($wsus.GetComputerTargetGroup($Computer.ComputerTargetGroupIds[($PatchGroup)].guid)).Name
	$Header = $WsusPatchGroup + "\" + $WsusComputerName
	$ServerOK=$true
	if ($WsusComputerName.indexof(".") -gt 0) {
			$svrshortname = $WsusComputerName.substring(0,($WsusComputerName.indexof(".")))
	} else {
			$svrshortname = $WsusComputerName
	}
	$reportline = "<TR><TD><a href=`"$wsuspatchgroup\$svrshortname.html`">$wsuspatchgroup\$svrshortname</a></TD><TD>$WsusComputerName</TD><TD>"+($WsusLastCheckin).ToString("dd-MM-yyyy HH:mm:ss") +"</TD><TD align=center>$WsusTotalUpdateCount</TD><TD align=center>$WsusReqPatchCount</TD><TD align=center>$WsusFailedPatchCount</TD>"
	If ($ExcludedSvrs -ne "") {
		ForEach($ExcludedSvr in $ExcludedSvrs) {
			If ($Header -match $ExcludedSvr) {
				$ServerOK=$false
				Write-Host $Header "Excluded" -ForegroundColor Red
				$reportline += "<TD align=center><font color=orange>Excluded</font></TD>"
			}
		}
	}
	If ($ServerOK) {
		If ($WsusPatchGroup -eq "Unassigned Computers") {
			$UnassignedComputers += $WsusComputerName
		}		
		$ReportFolder = $Cfg_BaseReportFolder + $WsusPatchGroup + "\"
		if (!(Test-Path $ReportFolder)) {
			New-Item $ReportFolder -type directory | out-null
		}
		$ServerDetails = $ReportFolder + $svrshortname + ".html"
		if ($WsusLastCheckin -lt (Get-Date).AddDays(-$Cfg_DaysSinceCheckIn) -or $WsusLastCheckin -eq "01/01/0001 00:00:00")
		{
			$Header = $WsusPatchGroup + "\" + $WsusComputerName + " >> Not Checked In in last " + $Cfg_DaysSinceCheckIn + " Days"
			write-host $Header -ForegroundColor Red
			"Report Generated $Today" | Out-File -FilePath $ServerDetails
			$Header | Out-File -FilePath $ServerDetails -Append
			$ComputersNotCheckingInCount += 1
			$ComputersNotCheckingIn += $WsusPatchGroup + "\" + $WsusComputerName + " (" + (get-date($WsusLastCheckin)).ToString("dd/MM/yyyy") + ")"
            $reportline += "<TD align=center><font color=red>Not Checking In</font></TD>"
			
			ForEach($UserDetail in $UsersFile) {
                If(($UserDetail.computer -eq $svrshortname)) {
                    $UserName = $UserDetail.name
                    $UserMail = $UserDetail.mail + $Cfg_Email_Domain
                    $UserBody = "Dear $UserName<br>Your computer $svrshortname has not checked into our Windows Updates server since "
                    $UserBody += (get-date($WsusLastCheckin)).ToString("dd/MM/yyyy")
                    $UserBody += ".<br><br>Please contact the service desk to correct this at your earliest convenience.<br>"
            
                    if ($Cfg_Email_User_Send) {
						$UserSmtp = New-Object System.Net.Mail.SmtpClient -argumentList $Cfg_Email_Server
						If($Cfg_Smtp_User) {
							$smtp.Credentials = New-Object System.Net.NetworkCredential -argumentList $Cfg_Smtp_User,$Cfg_Smtp_Password
						}
						$UserMessage = New-Object System.Net.Mail.MailMessage
						$UserMessage.From = New-Object System.Net.Mail.MailAddress($Cfg_Email_From_Address)
						$UserMessage.To.Add($UserMail)
						$UserMessage.Subject = $Cfg_Email_Subject_User + " $svrshortname - " + (get-date).ToString("dd/MM/yyyy")
						$UserMessage.isBodyHtml = $true
						$UserMessage.Body = $UserBody
						Write-Output "Sending Not Checked In e-mail to $UserMail"
						$UserSmtp.Send($UserMessage)
					}
                }
            }
		}
		Else
		{
			# Find all Updates in Update Scope for This Computer
			
			$Header = $WsusPatchGroup + "\" + $WsusComputerName + " >> $WsusTotalUpdateCount Total Updates >> $WsusReqPatchCount Required Updates >>"
			Write-Host $Header -ForegroundColor Yellow -NoNewline
			"Report Generated $Today <br>" | Out-File -FilePath $ServerDetails

			$NeededPatches = $Computer.GetUpdateInstallationInfoPerUpdate($UpdateScope)
			
			$UpdatesApproved = 0
			
			ForEach ($Patch in $NeededPatches) {
				#$PatchFound=0
				$PatchDetail = $wsus.GetUpdate($Patch.UpdateID)
				$PatchTitle = $PatchDetail.Title
                write-Debug $PatchTitle
				If ($Patch.UpdateApprovalAction -eq "Install") {
                    $UpdatesApproved +=1
                }
				"$PatchTitle - " + $Patch.UpdateApprovalAction + "<br>" | Out-File -FilePath $ServerDetails -Append
			}
			$line = "$Header $UpdatesApproved Approved Updates <br>"
			$line | Out-File -FilePath $ServerDetails -Append
			If ($UpdatesApproved -ge $Cfg_Unpatched_Threshold) {
				Write-Host " $UpdatesApproved Approved Updates" -ForegroundColor Red
				$ComputersNotPatching += $WsusPatchGroup + "\" + $WsusComputerName + " - " + $UpdatesApproved + " Approved Updates"
			} else {
				Write-Host " $UpdatesApproved Approved Updates" -ForegroundColor Yellow
			}
			#Write-Host "`n"
			$reportline += "<TD align=center>$UpdatesApproved</TD>"

            ForEach($UserDetail in $UsersFile) {
                If(($UserDetail.computer -eq $svrshortname) -and ($UpdatesApproved -ge $Cfg_Unpatched_Threshold)) {
                    $UserName = $UserDetail.name
                    $UserMail = $UserDetail.mail + $Cfg_Email_Domain
                    $UserBody = "Dear $UserName<br>Your computer $svrshortname currently has <font color=red> $UpdatesApproved </font> windows updates approved and pending installation.<br><br>"
                    $UserBody += "Please endeavour to install these updates at your earliest convenience."
                    
                    if ($Cfg_Email_User_Send) {
                        $UserSmtp = New-Object System.Net.Mail.SmtpClient -argumentList $Cfg_Email_Server
						If($Cfg_Smtp_User) {
							$smtp.Credentials = New-Object System.Net.NetworkCredential -argumentList $Cfg_Smtp_User,$Cfg_Smtp_Password
						}
                        $UserMessage = New-Object System.Net.Mail.MailMessage
                        $UserMessage.From = New-Object System.Net.Mail.MailAddress($Cfg_Email_From_Address)
                        $UserMessage.To.Add($UserMail)
                        $UserMessage.Subject = $Cfg_Email_Subject_User + " $svrshortname - " + (get-date).ToString("dd/MM/yyyy")
                        $UserMessage.isBodyHtml = $true
                        $UserMessage.Body = $UserBody
                        Write-Output "Sending Outstanding Updates e-mail to $UserMail"
                        $UserSmtp.Send($UserMessage)
                    }
                }
            }
		}
	}
    #Write-Host $reportline
	$reportline += "</TR>"
	$reportline | Out-file -FilePath $ReportFile -Append -encoding ASCII
    
}
$EndTime = (get-date).ToString("dd-MM-yyyy HH:mm:ss")
[string]$Report = "Start Time: " + $StartTime
Write-Host "`n" -NoNewline
Write-Host $Report
$Report + "<BR>" | Out-File -FilePath $Emailbody -encoding ASCII

[string]$Report = "End Time: " + $EndTime
Write-Host $Report
$Report + "<BR>" | Out-File -FilePath $Emailbody -Append -encoding ASCII

[string]$Report = $Computers.count 
$Report += " Computers Checked"
Write-Host "`n" -NoNewline
Write-Host $Report
"<BR>" | Out-File -FilePath $Emailbody -Append -encoding ASCII
$Report + "<BR>" | Out-File -FilePath $Emailbody -Append -encoding ASCII

"</TABLE></BODY>" | Out-file -FilePath $ReportFile -Append -encoding ASCII

Write-Host "Report URL: " $Cfg_Report_URL
$Cfg_Report_URL + "<BR>" | Out-File -FilePath $Emailbody -Append -encoding ASCII

if($ComputersNotCheckingInCount -gt 0) {
	[Array]::Sort([array]$ComputersNotCheckingIn)
	[string]$Report = $ComputersNotCheckingInCount
	$Report += " Computers Not Checked In in the last $Cfg_DaysSinceCheckIn days"
	Write-Host "`n" -NoNewline
	Write-Host $Report
	"<BR><U>" | Out-File -FilePath $Emailbody -Append -encoding ASCII
	$Report + "</U><BR>" | Out-File -FilePath $Emailbody -Append -encoding ASCII
	ForEach ($ComputerNotCheckingIn In $ComputersNotCheckingIn) {
		$ComputerNotCheckingIn + "<BR>" | Out-File -FilePath $Emailbody -Append -encoding ASCII
	}
}

if($ComputersNotPatching.Count -gt 0) {
	[Array]::Sort([array]$ComputersNotPatching)
	[string]$Report = $ComputersNotPatching.Count
	$Report += " Computers with $Cfg_Unpatched_Threshold or more approved updates outstanding"
	Write-Host "`n" -NoNewline
	Write-Host $Report
	"<BR><U>" | Out-File -FilePath $Emailbody -Append -encoding ASCII
	$Report + "</U><BR>" | Out-File -FilePath $Emailbody -Append -encoding ASCII
	ForEach ($ComputerNotPatching In $ComputersNotPatching) {
		$ComputerNotPatching + "<BR>" | Out-File -FilePath $Emailbody -Append -encoding ASCII
	}
}

"<br>`nScript Details: "+ ($MyInvocation.MyCommand.Definition) +" on "+ $env:COMPUTERNAME +"<br>`n" | Out-File -FilePath $Emailbody -Append -encoding ASCII

if ($Cfg_Email_Report_Send) {
    $Body = Get-Content $Emailbody
    Write-Host "`nSending Report E-Mail to" $Cfg_Email_To_Address
    $smtp = New-Object System.Net.Mail.SmtpClient -argumentList $Cfg_Email_Server
	If($Cfg_Smtp_User) {
		$smtp.Credentials = New-Object System.Net.NetworkCredential -argumentList $Cfg_Smtp_User,$Cfg_Smtp_Password
	}
    $message = New-Object System.Net.Mail.MailMessage
    $message.From = New-Object System.Net.Mail.MailAddress($Cfg_Email_From_Address)
    $message.To.Add($Cfg_Email_To_Address)
    $message.Subject = $Cfg_Email_Subject + " - " + (get-date).ToString("dd/MM/yyyy")
    $message.isBodyHtml = $true
    $message.Body = $Body
    $attachment = new-object Net.Mail.Attachment($ReportFile) 
    $message.Attachments.Add($attachment)
    $smtp.Send($message)
}
