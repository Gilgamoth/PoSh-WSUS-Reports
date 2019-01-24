# WSUS Report - Config File
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

$Cfg_Version = "2019-01-24-0120"

# **************************** CONFIG SECTION ****************************

# Only needed if not localhost or not default port, leave blank otherwise.
	$Cfg_WSUSServer = "WSUSServer.FQDN.local"
	$Cfg_WSUSSSL = $true
	$Cfg_WSUSPort = 8531

# Report URL on IIS to see reports remotely - Remember to put the trailing slash in
	$Cfg_BaseReportFolder = "C:\Program Files\Update Services\WebServices\Root\reports\"

# Servers Excluded from Patching Report
	$Cfg_ExcludedServerFile = "c:\scripts\excludedsvrs.txt"

# User List
	$Cfg_UsersFile = "c:\scripts\users.csv"

# Targets a specific Computer Group in WSUS or leave blank for all
	$Cfg_WSUSTarget = ""

# How long before server should be listed as AWOL
	$Cfg_DaysSinceCheckIn = 17

# How many patches unapplied before flagged as an issue
	$Cfg_Unpatched_Threshold = 1

# E-Mail Report Details
	$Cfg_Email_To_Address = "Patching.Team@domain.local"
	$Cfg_Email_From_Address = "Patching.Manager@domain.local"
	$Cfg_Email_Subject = "WSUS: Outstanding Patches Report"
	$Cfg_Email_Subject_User = "Windows Updates For Your Machine"
	$Cfg_Email_Server = "mail.domain.local"
	$Cfg_Email_Domain = "@domain.local"
    $Cfg_Email_Report_Send = $true
    #$Cfg_Email_Report_Send = $false
    $Cfg_Email_User_Send = $true
    #$Cfg_Email_User_Send = $false

# E-Mail Server Credentials to send report (Leave Blank if not Required)
	$Cfg_Smtp_User = ""
	$Cfg_Smtp_Password = ""

# Report URL on IIS to see reports
	IF($Cfg_WSUSSSL) {
		$Cfg_Report_URL="https://"+$Cfg_WSUSServer+":"+$Cfg_WSUSPort+"/reports/"
	} else {
		$Cfg_Report_URL="http://"+$Cfg_WSUSServer+":"+$Cfg_WSUSPort+"/reports/"
	}
	
