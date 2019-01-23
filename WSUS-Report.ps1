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

$App_Version = "2019-01-23-2200"

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
$ReportFile = $Cfg_BaseReportFolder + "report.csv"
$Emailbody =  $Cfg_BaseReportFolder + "emailbody.txt"
$Today=(get-date).ToString("yyyy-MM-dd")
$UnapprovedUpdateCount = 0
$UnapprovedUpdates = @()
$NewPatchesFound = 0
$NewPatches = @()
$UnassignedComputers = @()
$ComputersNotPatching = @()
$ExcludedSvrs = ""
$wsus = ""

"Company, ServerName, FQDN, LastCheckIn, TotalPatches, FailedPatches, ApprovedPatches" | Out-file -FilePath $ReportFile

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

$Computers = $wsus.GetComputerTargets($ComputerScope)

# Set Updates Scope to Critical & Security Updates for ForeFront and Windows

$UpdateScope = new-object Microsoft.UpdateServices.Administration.UpdateScope
$UpdateScope.IncludedInstallationStates = [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::NotInstalled, [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::Downloaded, [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::Failed
#$wsus.GetUpdateClassifications() | where {"Security Updates","Critical Updates" -contains $_.Title} | %{$UpdateScope.Classifications.Add($_)} | out-null
#$wsus.GetUpdateCategories() | where {"Windows","Forefront" -contains $_.Title} | %{$UpdateScope.Categories.Add($_)} | out-null

# Process for each computer in Computer Scope

ForEach($Computer in $Computers) {
	$WsusComputerName = $Computer.FullDomainName
	$WsusTotalPatchCount = $Computer.GetUpdateInstallationSummary()
	$WsusFailedPatchCount = $WsusTotalPatchCount.FailedCount
	$WsusReqPatchCount = $Computer.GetUpdateInstallationSummary($UpdateScope)
	$WsusLastCheckin = $Computer.LastReportedStatusTime
	$WsusCheckinServer = $Computer.UpdateServer.ServerName
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
	$reportline = "$wsuspatchgroup, $svrshortname, $WsusComputerName, $WsusLastCheckin, $WsusTotalUpdateCount, $WsusFailedPatchCount,"
	If ($ExcludedSvrs -ne "") {
		ForEach($ExcludedSvr in $ExcludedSvrs) {
			If ($Header -match $ExcludedSvr) {
				$ServerOK=$false
				Write-Host $Header "Excluded" -ForegroundColor Red
				$reportline += "Excluded"
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
		$ServerDetails = $ReportFolder + $svrshortname + ".txt"
		if ($WsusLastCheckin -lt (Get-Date).AddDays(-$Cfg_DaysSinceCheckIn) -or $WsusLastCheckin -eq "01/01/0001 00:00:00")
		{
			$Header = $WsusPatchGroup + "\" + $WsusComputerName + " >> Not Checked In in last " + $Cfg_DaysSinceCheckIn + " Days"
			write-host $Header -ForegroundColor Red
			(get-date).ToString("dd-MM-yyyy HH:mm") + "<BR>" | Out-File -FilePath $ServerDetails
			$Header + "<BR>" | Out-File -FilePath $ServerDetails -Append
			$ComputersNotCheckingInCount += 1
			$ComputersNotCheckingIn += $WsusPatchGroup + "\" + $WsusComputerName + " (" + (get-date($WsusLastCheckin)).ToString("dd/MM/yyyy") + ")"
            ForEach($UserDetail in $UsersFile) {
                If(($UserDetail.computer -eq $svrshortname)) {
                    $UserName = $UserDetail.name
                    $UserMail = $UserDetail.mail + $Cfg_Email_Domain
                    $UserBody = "Dear $UserName<br>Your computer $svrshortname has not checked into our Windows Updates server since "
                    $UserBody += (get-date($WsusLastCheckin)).ToString("dd/MM/yyyy")
                    $UserBody += ".<br><br>Please contact the service desk to correct this at your earliest convenience.<br>"
            
                    Write-Host "`nSending E-Mail to" $UserMail
                    $UserSmtp = New-Object System.Net.Mail.SmtpClient -argumentList $Cfg_Email_Server
                    #$smtp.Credentials = New-Object System.Net.NetworkCredential -argumentList $Cfg_Smtp_User,$Cfg_Smtp_Password
                    $UserMessage = New-Object System.Net.Mail.MailMessage
                    $UserMessage.From = New-Object System.Net.Mail.MailAddress($Cfg_Email_From_Address)
                    $UserMessage.To.Add($UserMail)
                    $UserMessage.Subject = $Cfg_Email_Subject_User + " - " + (get-date).ToString("dd/MM/yyyy")
                    $UserMessage.isBodyHtml = $true
                    $UserMessage.Body = $UserBody
                    Write-Output "Sending Not Checked In e-mail to $UserMail"
                    $UserSmtp.Send($UserMessage)
                }
            }
		}
		Else
		{
			# Find all Updates in Update Scope for This Computer
			
			$Header = $WsusPatchGroup + "\" + $WsusComputerName + " >> " + $WsusTotalUpdateCount + " Total Updates >> " + $WsusFailedPatchCount + " Failed Patches "
			Write-Host $Header -ForegroundColor Yellow -NoNewline
			(get-date).ToString("dd-MM-yyyy HH:mm") + "<BR>" | Out-File -FilePath $ServerDetails
			$Header + "<BR>" | Out-File -FilePath $ServerDetails -Append

			$NeededPatches = $Computer.GetUpdateInstallationInfoPerUpdate($UpdateScope)
			
			$UpdatesApproved = 0
			$UpdatesDeclined = 0
			$UpdatesUnapproved = 0
			
			ForEach ($Patch in $NeededPatches) {
				$PatchFound=0
				$PatchDetail = $wsus.GetUpdate($Patch.UpdateID)
				$PatchTitle = $PatchDetail.Title
                write-Debug $PatchTitle
				$first=$PatchTitle.IndexOf("(")
                if($first -gt 0) {
				    $second = ($PatchTitle.Substring($first+1)).indexof(")")
				    $kbnum = $PatchTitle.Substring(($first+1),($second))
				    $kbtitle = $PatchTitle.Substring(0,($first-1))
				    $line = $kbnum + "`t" + $PatchTitle.Substring(0,$First-1)
                } else {
                    $kbnum = "N/A"
                }
				If ($Patch.UpdateApprovalAction -eq "Install") {
                    $UpdatesApproved +=1
                    $kbtitle = $PatchTitle
                    $line = $kbnum + "`t" + $PatchTitle
                }
			}
			$line = "" + $UpdatesApproved + " Approved Updates"
			#$line +=  $UpdatesDeclined + " Declined Updates "  + $UpdatesUnapproved + " Unapproved Updates"
			#write-host $line
			$line | Out-File -FilePath $ServerDetails -Append
			If ($UpdatesApproved -ge $Cfg_Unpatched_Threshold) {
				Write-Host ">> $UpdatesApproved Updates Approved" -ForegroundColor Red
				$ComputersNotPatching += $WsusPatchGroup + "\" + $WsusComputerName + " - " + $UpdatesApproved + " Approved Updates"
			} else {
				Write-Host ">> $UpdatesApproved Updates Approved" -ForegroundColor Yellow
			}
			#Write-Host "`n"
			$reportline += $UpdatesApproved

            ForEach($UserDetail in $UsersFile) {
                If(($UserDetail.computer -eq $svrshortname) -and ($UpdatesApproved -gt 1)) {
                    $UserName = $UserDetail.name
                    $UserMail = $UserDetail.mail + $Cfg_Email_Domain
                    $UserBody = "Dear $UserName<br>Your computer $svrshortname currently has <font color=red> $UpdatesApproved </font> windows updates approved and pending installation.<br><br>"
                    $UserBody += "Please endeavour to install these updates at your earliest convenience."
                    
                    if ($Cfg_Email_Send) {
                        Write-Host "`nSending E-Mail to" $UserMail
                        $UserSmtp = New-Object System.Net.Mail.SmtpClient -argumentList $Cfg_Email_Server
                        #$smtp.Credentials = New-Object System.Net.NetworkCredential -argumentList $Cfg_Smtp_User,$Cfg_Smtp_Password
                        $UserMessage = New-Object System.Net.Mail.MailMessage
                        $UserMessage.From = New-Object System.Net.Mail.MailAddress($Cfg_Email_From_Address)
                        $UserMessage.To.Add($UserMail)
                        $UserMessage.Subject = $Cfg_Email_Subject_User + " - " + (get-date).ToString("dd/MM/yyyy")
                        $UserMessage.isBodyHtml = $true
                        $UserMessage.Body = $UserBody
                        Write-Output "Sending Outstanding Patches e-mail to $UserMail"
                        $UserSmtp.Send($UserMessage)
                    }
                }
            }
		}
	}
    #Write-Host $reportline
	$reportline | Out-file -FilePath $ReportFile -Append
    
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
	$Report += " Computers with $Cfg_Unpatched_Threshold or more approved patches outstanding"
	Write-Host "`n" -NoNewline
	Write-Host $Report
	"<BR><U>" | Out-File -FilePath $Emailbody -Append -encoding ASCII
	$Report + "</U><BR>" | Out-File -FilePath $Emailbody -Append -encoding ASCII
	ForEach ($ComputerNotPatching In $ComputersNotPatching) {
		$ComputerNotPatching + "<BR>" | Out-File -FilePath $Emailbody -Append -encoding ASCII
	}
}

if ($Cfg_Email_Send) {
    $Body = Get-Content $Emailbody
    Write-Host "`nSending Report E-Mail to" $Cfg_Email_To_Address
    $smtp = New-Object System.Net.Mail.SmtpClient -argumentList $Cfg_Email_Server
    #$smtp.Credentials = New-Object System.Net.NetworkCredential -argumentList $Cfg_Smtp_User,$Cfg_Smtp_Password
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
