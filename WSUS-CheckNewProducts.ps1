# Windows Update Patch Approval Script
# Copyright (C) 2018 Steve Lunn (gilgamoth@gmail.com)
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

$App_Version = "2022-07-26-1041"

Clear-Host
Set-PSDebug -strict
$ErrorActionPreference = "SilentlyContinue"
[GC]::Collect()

# **************************** VARIABLE START ******************************

# Only needed if not localhost or not default port, leave blank otherwise.
    $Cfg_WSUSServer = "updates.wolftech.org.uk" # WSUS Server Name
    $Cfg_WSUSSSL = $true # WSUS Server Using SSL
    $Cfg_WSUSPort = 8531 # WSUS Port Number

    #$Cfg_WSUSServer = "wolfpas02" # WSUS Server Name
    #$Cfg_WSUSSSL = $false # WSUS Server Using SSL
    #$Cfg_WSUSPort = 8530 # WSUS Port Number

# Required Variables
    $Cfg_ReportDays = -7
    $Cfg_ReportDate = (Get-Date).AddDays($Cfg_ReportDays)

# E-Mail Report Details
	$Cfg_Email_To_Address = "stevel@wolftech.local"
	$Cfg_Email_From_Address = "WSUS-Report@wolftech.local"
	$Cfg_Email_Subject = "WSUS: New Classifications & Products " + $env:computername
	$Cfg_Email_Server = "wolfpex01"
    $Cfg_Email_Send = $false
    $Cfg_Email_Send = $true # Comment out if no e-mail required
	[string]$Email_Body = "Getting New Classifications & Products since " + $Cfg_ReportDate.ToString("yyyy-MM-dd") +" for server " + $env:computername +" on URL "+ $Cfg_WSUSServer +":"+ $Cfg_WSUSPort +"<br>`n<br>`n"

# E-Mail Server Credentials to send report (Leave Blank if not Required)
	$Cfg_Smtp_User = ""
	$Cfg_Smtp_Password = ""

$StartTime = Get-Date

$Email_Body += "<TABLE BORDER=1><TR><TH WIDTH=150>Type</TH><TH WIDTH=500>Title</TH><TH WIDTH=175>Arrival Date</TH></TR>"
# Get new classifications
Write-Host "Getting New Classifications"
#(Get-WsusServer).GetUpdateClassifications($Cfg_ReportDate,$StartTime) | Select Title, Description, ReleaseNotes, ArrivalDate
$New_Classifications = (Get-WsusServer).GetUpdateClassifications($Cfg_ReportDate,$StartTime)
If ($New_Classifications.count -gt 0) {
	ForEach ($Classifications in $New_Classifications) {
		$LineData = "<TR><TD>Classification</TD><TD>"+$Classifications.Title+"</TD><TD ALIGN=CENTER>"+$Classifications.ArrivalDate+"</TD></TR>"
		write-host $Classifications.Title $Classifications.ArrivalDate
        $Email_Body += $LineData
    }
} else {
	$LineData = "<TR><TD COLSPAN=3>No New Classifications Found</TD></TR>"
	write-host " - No new Classifications Found"
    $Email_Body += $LineData

}
 
# Get what new products were added over the month
Write-Host "Getting New Products"
#(Get-WsusServer).GetUpdateCategories($Cfg_ReportDate,$StartTime) |  Select Title, Id, ArrivalDate | ft -AutoSize
$New_Products = (Get-WsusServer).GetUpdateCategories($Cfg_ReportDate,$StartTime)
If ($New_Products.count -gt 0) {
	ForEach ($Product in $New_Products) {
		$LineData = "<TR><TD>Product</TD><TD>"+$Product.Title+"</TD><TD ALIGN=CENTER>"+$Product.ArrivalDate+"</TD></TR>"
		write-host $Product.Title $Product.ArrivalDate
        $Email_Body += $LineData
    }
} else {
	$LineData = "<TR><TD COLSPAN=3>No New Products Found</TD></TR>"
	write-host " - No new Products Found"
    $Email_Body += $LineData

}

$Email_Body += "</TABLE><BR>`n"

$EndTime = Get-Date
$RunTime = $EndTime - $StartTime
$FormatTime = "{0:N2}" -f $RunTime.TotalMinutes

Write-Host "`nStart Time: $StartTime"
Write-Host "End Time: $EndTime"

$Email_Body += "<br>`nScript Details: "+ ($MyInvocation.MyCommand.Definition) +" on "+ $env:COMPUTERNAME +"<br>`n"
$Email_Body += "<br>Started at $StartTime<br>"
$Email_Body += "Finished at $EndTime<br>"
$Email_Body += "Job took $FormatTime minutes to run<br>"

if ($Cfg_Email_Send) {
    $smtp = New-Object System.Net.Mail.SmtpClient -argumentList $Cfg_Email_Server
    if ($Cfg_Smtp_User) {
        $smtp.Credentials = New-Object System.Net.NetworkCredential -argumentList $Cfg_Smtp_User,$Cfg_Smtp_Password
    }
    $message = New-Object System.Net.Mail.MailMessage
    $message.From = New-Object System.Net.Mail.MailAddress($Cfg_Email_From_Address)
    $message.To.Add($Cfg_Email_To_Address)
    $message.Subject = $Cfg_Email_Subject + " " + (get-date).ToString("dd/MM/yyyy")
    $message.isBodyHtml = $true
    $message.Body = $Email_Body
    Write-Host "`nSending Report E-Mail to" $message.To.address
    $smtp.Send($message)
}
