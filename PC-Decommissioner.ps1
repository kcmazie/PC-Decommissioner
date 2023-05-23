Param(
    [String]$TargetPC,
    [String]$Technician,    
    [Switch]$Debug = $False,
    [Switch]$Console = $False,
    [Switch]$SendEmail = $True
    )
<#==============================================================================
          File Name : PC-Decommissioner.ps1
    Original Author : Kenneth C. Mazie (kcmjr AT kcmjr DOT com)
                    :
        Description : Detects DNS assigned to local system. Pulls all records from selected DNS server and performs
                    : both a ping and an NSlookup on the record. Records to Excel and/or Console. Use to validate
                    : existing DNS records.
                    :
              Notes : Normal operation is with no command line options except as noted below.
                    :
  Optional arguments: -Debug $true (defaults to false). Sends email to debug address from config file.
                    : -Console $true (defaults to false). Displays runtime info on Console
                    : -SendEmail $False (defaults to true). Will stop emails from going out.
                    : -Technician (REQUIRED) The AD user ID of the person submitting the report. Will fail if invalid.
                    : -Target (REQUIRED) The name of the system to delete.
                    :
           Warnings : None
                    :
              Legal : Public Domain. Modify and redistribute freely. No rights reserved.
                    : SCRIPT PROVIDED "AS IS" WITHOUT WARRANTIES OR GUARANTEES OF
                    : ANY KIND. USE AT YOUR OWN RISK. NO TECHNICAL SUPPORT PROVIDED.
                    : That being said... If you find bugs, PLEASE let me know so I can correct them.
                    :
            Credits : Code snippets and/or ideas came from many sources including but
                    : not limited to the following:
                    :
     Last Update by : Kenneth C. Mazie
    Version History : v1.00 - 10-02-18 - Original
     Change History : v2.00 - 00-00-00 -
                    #>
     $ScriptVer = "1.00"
                    <#
                    :
#===============================================================================#>
<#PSScriptInfo
.VERSION 1.00
.AUTHOR Kenneth C. Mazie (kcmjr AT kcmjr DOT com)
.DESCRIPTION
Accepts a PC name and a user name as command line inputs and then removes the designated PC from
Active Directory and DNS both. Validates the user before execution. Emails an HTML report
to designated recipients.
#>
#requires -version 5.1

Clear-host    

#==[ For Testing ]#=============================
#$TargetPC = "delete-test"
#$Technician = "testuser"
#$Script:Debug = $false
#$Script:Console = $true
#$SendEmail = $false
#$ByPassList = "" #("delete-test")
#===============================================

If ($Debug){$Script:Debug = $true}    
If ($Console){$Script:Console = $true}    
If ($SendEmail){$Script:SendEmail = $true}    

$DNSDomainName = (Get-ADDomain).DNSroot  
$ErrorActionPreference = "silentlycontinue"
$ScriptName = ($MyInvocation.MyCommand.Name).split(".")[0] 
$ConfigFile = $PSScriptRoot+'\'+$ScriptName+'.xml'
$HostName = ""
$Result = ""

Function LoadConfiguration{         
    #--[ Read and load configuration file ]-------------------------------------
    If (!(Test-Path $ConfigFile)){                                               #--[ Error out if configuration file doesn't exist ]--
        $Script:MessageBody = "MISSING CONFIG FILE. Script aborted."
        Write-Host "CONFIGURATION FILE NOT FOUND - EXITING" -ForegroundColor Red 
        break
    }Else{
        Try{
            [xml]$Configuration = Get-Content $ConfigFile -ErrorAction "stop"                    #--[ Read & Load XML ]--
        }Catch{
            $Script:MessageBody = $_.Exception.Message + " - Script aborted."
            Write-Host $_.Exception.Message " - Script aborted." -ForegroundColor Red 
            break
        }
        #$DnsServer = $Configuration.Settings.General.DnsServer #--[ Detected. See below ]--
        $Script:DnsDomain = $Configuration.Settings.General.Domain              #--[ Detected. See below ]--
        $Script:DebugEmail = $Configuration.Settings.Email.Debug 
        $Script:DebugTarget = $Configuration.Settings.General.DebugTarget 
        $Script:ValidGroup = $Configuration.Settings.General.ValidGroup
        $Script:eMailRecipient = $Configuration.Settings.Email.To
        $Script:eMailFrom = $Configuration.Settings.Email.From    
        $Script:eMailHTML = $Configuration.Settings.Email.HTML
        $Script:eMailSubject = $Configuration.Settings.Email.Subject
        $Script:SmtpServer = $Configuration.Settings.Email.SmtpServer
        $Script:UserName = $Configuration.Settings.Credentials.Username
        $EncryptedPW = $Configuration.Settings.Credentials.Password
        $Base64String = $Configuration.Settings.Credentials.Key
        #$ReportName = $Configuration.Settings.General.ReportName
        $Script:BypassList = ($Configuration.Settings.Bypass.List).Split(',')
        $ByteArray = [System.Convert]::FromBase64String($Base64String)
        $Script:Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName, ($EncryptedPW | ConvertTo-SecureString -Key $ByteArray)
        #$Script:Credential.password #--[ Will expose the encrypted password. Use with caution ]--
    }

}

Function SendMessage {      #--[ Send Email Report ]--------------------------------------------
    $Email = $null
    $Email = New-Object System.Net.Mail.MailMessage
    $Email.From = $Script:EmailFrom
    $Email.Subject = $Script:eMailSubject
    $Email.IsBodyHtml = $Script:eMailHTML
    $Email.Body += $Script:MessageBody    
    If ($Script:SendEmail){
        If ($Script:Debug){
            $Email.To.Add($Script:DebugEmail)                                                  #--[ Debug destination email address ]--
            If ($Script:Console){write-host "`n--[ DEBUG Email sent ]--" -ForegroundColor Green}
        }Else{    
            $Email.To.Add($eMailRecipient)                                              #--[ Destination email address ]--
            If ($Script:Console){write-host "`n--[ Email sent ]--" -ForegroundColor Green}
        }
        $SMTP = new-object Net.Mail.SmtpClient($Script:SMTPServer)
        $SMTP.Send($Email)
    }Else{
        If ($Script:Console){write-host "`n--[ Email Disabled. Nothing Sent ]--" -ForegroundColor Red}
    }
}      

Function ConsoleColor {  #--[ Detect Console color and adjust accordingly ]------------------------------
    If ((Get-Host).UI.RawUI.BackgroundColor -eq "White"){
        $Script:FgGreen = "DarkGreen"
        $Script:FgRed = "DarkRed"
        $Script:FgYellow = "DarkCyan"
        $Script:FgBlue = "DarkCyan"
        $Script:FgCyan = "DarkCyan"
        $Script:FgMagenta = "DarkCyan"
        $Script:FgGray = "DarkGray"
        $Script:FgText = "Black"
    }Else{
        $Script:FgGreen = 'Green'
        $Script:FgRed = 'Red'
        $Script:FgYellow = 'Yellow'
        $Script:FgBlue = 'Blue'
        $Script:FgCyan = 'Cyan'
        $Script:FgMagenta = 'Magenta'
        $Script:FgGray = 'Gray'
        $Script:FgText = 'White'
    }
}

#--End of Functions ]-----------------------------------------------------------
 
#==[ Main Process ]=============================================================
Try{Get-PSSession | Remove-PSSession}Catch{
    If ($Script:Console){Write-host $_.Exception.Message -ForegroundColor Yellow}
}
ConsoleColor 
LoadConfiguration

$ErrorActionPreference = "stop"
Try{
    $Technician = Get-ADUser -Identity $Technician -Credential $Script:Credential -Properties * #GivenName,Surname,memberof,DisplayName -ErrorAction "stop"
   # $Technician
}Catch{
    $Script:MessageBody = $_.Exception.Message
    Write-Host $Script:MessageBody -ForegroundColor Red
    $Script:MessageBody += "<br>Script cannot run without a valid user - Exiting"
    SendMessage
    Break
}    

#--[ Verify permission to run ]-------------------
if ($Technician.memberof -like "*$Script:ValidGroup*"){
    #--[ User is a member of appropriate AD group ]--
    If ($Script:Debug){Write-Host "--[ User is a member of appropriate AD group ]--" -ForegroundColor green}
}Else{
    $Script:MessageBody = "User "+($Technician.DisplayName)+" does not have appropriate AD group membership to execute this operation"
    Write-Host $Script:MessageBody -ForegroundColor Red
    $Script:MessageBody += "<br>Process cannot run without a validated user - Exiting"
    SendMessage
    Break
}
    
<#--[ Reference Only ]--------------------------
$Technician.Name
$Technician.GivenName
$Technician.UserPrincipalName
$Technician.SamAccountName
$Technician.Enabled
$Technician.memberof
#----------------------------------------------#>

#--[ Identify DNS servers in the domain ]--------------------------
$DNSServer = ""
$DNSServerList = ""

#--[ Detect domain unless one is specified in config file. ]--
if ([string]::IsNullOrEmpty($DnsDomain)){   
    $DNSDomainName = (Get-ADDomain).DNSRoot
}Else{
    $DnsDomainName = $DnsDomain
} 

$Lookup = Invoke-Command -ScriptBlock {nltest.exe /dnsgetdc:$DNSDomainName} 
$DNSServerList = ($Lookup[2..$($Lookup.Count - 2)]).Trim() 

#--[ Get authentication DC ]--------------------------------------
$LogonServer = ($ENV:LOGONSERVER).trimstart("\").ToLower()
 
#--[ Get a random DNS server that is NOT the authenticating DC ]------------------
$RndDNS = $LogonServer
while ($RndDNS.Split(".")[0] -eq $LogonServer){
    $RndDNS = Get-Random -InputObject $DNSServerList -Count "1" # | where $_ -NotMatch $LogonServer
}
$DNSServer = (($RndDNS.Split(".")[0]).Trim()).ToString()
$DNSServerIP = $RndDNS.Split(".")[1]
#--[ Selecting a random DC like this for DNS processing avoids the issue with double logon to the same DC ]--

#--[ Use for manual selection of one of the local system DNS servers from IP settings ]--
#$DNS1 = (Get-DnsClientServerAddress -AddressFamily "IPv4").ServerAddresses[0]
#$DNS2 = (Get-DnsClientServerAddress -AddressFamily "IPv4").ServerAddresses[1]
#$DNSServerIP = $DNS2
#$DNSServer = [System.Net.Dns]::GetHostByAddress($DNSServerIP).Hostname

$PropertyList = @(
    "HostName",
    "Technician",
    "TechValid",
    "Date",
    "ADStatus",
    "DomainName",
    "Distinguishedname",
    "DNSServer",
    "ObjectClass",
    "ObjectGUID",
    "SamAccountName",
    "SID",
    "AdAction",
    "AdResult",
    "DnsDetect",
    "DnsSession",
    "DnsAction",
    "DnsResult",
    "Status",
    "Result"
)

$TargetRecordObj = New-Object -TypeName PSObject   #--[ Create a new object to hold results, pre-populate with blanks ]--
ForEach ($Property in $PropertyList){
    Add-Member -InputObject $TargetRecordObj -MemberType NoteProperty -Name $Property -Value "" -ErrorAction "stop" -Force
}

If ($Script:Console){write-host "--[ Processing $TargetPC ]--"`n -ForegroundColor Yellow}
$TargetRecordObj.HostName = ($TargetPC).ToUpper() 
$TargetRecordObj.Technician = $Technician.DisplayName
$TargetRecordObj.TechValid = "Validated" 
$TargetRecordObj.Date = ([System.DateTime]::Now) 
$TargetRecordObj.DomainName = $DNSDomainName
$TargetRecordObj.DNSServer = ($DNSServer.ToUpper())

Try{  #--[ Test for target System ]--
    $Lookup = Get-ADComputer -Identity $TargetPC -Credential $Script:Credential 
    If ($Lookup.Enabled = "true"){
        $TargetRecordObj.ADStatus = "Enabled"
    }Else{
        $TargetRecordObj.ADStatus = $Lookup.Enabled
    }
    $TargetRecordObj.HostName = ($Lookup.Name).ToUpper()
    $TargetRecordObj.DistinguishedName = $Lookup.DistinguishedName 
    $TargetRecordObj.ObjectClass = ($Lookup.ObjectClass).ToUpper()
    $TargetRecordObj.ObjectGUID = $Lookup.ObjectGUID
    $TargetRecordObj.SamAccountName = $Lookup.SamAccountName
    $TargetRecordObj.SID = $Lookup.SID
}Catch{
    $TargetRecordObj.ADStatus = $_.Exception.Message
    If ($Script:Console){$_.Exception.Message}
}

If ($BypassList -contains $TargetRecordObj.Hostname) {    
    $TargetRecordObj.ADStatus = "On Bypass List"
}Else{
    #--[ Kill Active Directory Record ]-----------------------------------------------------------
    If ($Debug){
        $TargetRecordObj.AdAction = "Simulating Delete"
        Remove-ADComputer -Identity $TargetPC -Credential $Script:Credential -whatif '2>&1' | out-null    #--[ Simulate deletion ]--
    }Else{   
        If ($TargetRecordObj.ADStatus -like "*Cannot find*"){     
            $TargetRecordObj.AdAction = "No Action"
        }Else{    
            Try{
                Remove-ADComputer -Identity $TargetPC -Credential $Script:Credential -Confirm:$false | out-null    #--[ Attempt deletion ]--
                $TargetRecordObj.AdAction = "Deleting AD record"
            }Catch{
                $TargetRecordObj.AdResult = $_.Exception.Message   
            }    
            Start-Sleep -Seconds 1
            Try{    #--[ Verify AD Deletion ]--
                $Result = Get-ADComputer -Identity $TargetPC -Credential $Script:Credential   #--[ Verify deletion ]--
                $TargetRecordObj.AdResult = "Ad Deletion FAILED"  
            }Catch{            
                If ($_.Exception.Message -like "*Cannot find an object*"){
                    $TargetRecordObj.AdResult = "Deletion verified"                           
                }Else{
                    $TargetRecordObj.AdResult = $_.Exception.Message
                }
            } 
        }
    }
    
    If ($Debug){  #--[ Simulate DNS by just detecting the record ]------------------------------
        $TargetRecordObj.DnsAction = "Simulating Delete"
        Try {   #--[ Check DNS for a record ]--
            $Check = cmd /c nslookup.exe $TargetRecordObj.Hostname $DNSServer '2>&1' | out-string
            If ($Check -Like "*Non-existent domain*"){
                $TargetRecordObj.DnsDetect = "Not Found"
            }Else{
                $TargetRecordObj.DnsDetect = "Record Found"
            }                            
        }Catch{
            $TargetRecordObj.DnsDetect = $_.Exception.Message                            
        }
    }Else{    
        #--[ Kill DNS Record ]-------------------------------------------------------------------
        Try {   #--[ Check DNS for a record ]--
            $Check = cmd /c nslookup.exe $TargetRecordObj.Hostname $DNSServer '2>&1' | out-string
            If ($Check -Like "*Non-existent domain*"){
                $TargetRecordObj.DnsDetect = "Not Found"
            }Else{
                $TargetRecordObj.DnsDetect = "Record Found"
            }                            
        }Catch{
            $TargetRecordObj.DnsDetect = $_.Exception.Message                            
        }

        If ($TargetRecordObj.DnsDetect -like "*Record Found*"){
            #--[ Open a PS session to selected DC/DNS server ]---------------
            Try{       
                $Session = New-PSSession -ComputerName $DNSServer -Credential $Script:Credential -ErrorAction "stop"
                $TargetRecordObj.DnsSession = $Session.State
            }Catch{
                $Msg = "PS Session to $DNSServer has failed... Script aborted. "+$_.Exception.Message
                $TargetRecordObj.DnsSession = $Msg
                Break
            }

             #--[ Enter the session ]-------------------------
            Try{      
                $TargetRecordObj.DnsAction = "Attempting Record Purge"
                Enter-PSSession -Session $Session -ErrorAction "stop"
                Invoke-Command -Session $Session -ScriptBlock {
                    $Lookup = Remove-DnsServerResourceRecord -ZoneName $Using:TargetRecordObj.DomainName -RRType A -Name $Using:TargetRecordObj.Hostname -ComputerName $Using:TargetRecordObj.DNSServer -Confirm:$false -Force
                    #Get-WmiObject -namespace "root\MicrosoftDNS" -Class MicrosoftDNS_$Using:Type -ComputerName "$Using:DNSServer" -Filter "IPAddress = '$Using:TargetRecordObj.NSLookupResult_Target_IP'" -ErrorAction "Stop" | Remove-WmiObject -ErrorAction "stop"
                    Return $Lookup
                } 
            }Catch{
                $TargetRecordObj.DnsResult = $_.Exception.Message                          
            }   
            Exit-PSSession     

            #--[ Verify DNS record Deletion ]--------------------
            Try {
                $ReCheck = cmd /c nslookup.exe $TargetRecordObj.Hostname $DNSServer '2>&1' | out-string
                If ($ReCheck -Like "*Non-existent domain*"){
                    $TargetRecordObj.DnsResult = "Record Purge Verified"
                }Else{
                    $TargetRecordObj.DnsResult = "DNS Record Still Exists"
                }                            
            }Catch{
                $TargetRecordObj.DnsResult = $_.Exception.Message                            
            }
        }Else{
            $TargetRecordObj.DnsResult = "No Action"
        }  
    }  
}
$TargetRecordObj.Result = $Result

#--[ Create header for email html content ]------------------
$Script:MessageBody = @() 
$Script:MessageBody += '
<style Type="text/css">
    table.myTable { border:5px solid black;border-collapse:collapse; }
    table.myTable td { border:2px solid black;padding:5px}
    table.myTable th { border:2px solid black;padding:5px;background: #949494 }
    table.bottomBorder { border-collapse:collapse; }
    table.bottomBorder td, table.bottomBorder th { border-bottom:1px dotted black;padding:5px; }
    tr.noBorder td {border: 0; }
</style>'

$Script:MessageBody += 
'<table class="myTable">
<tr class="noBorder"><td colspan=2><center><h1>- ' + $eMailSubject + ' -</h1></td></tr>
<tr class="noBorder"><td colspan=2><center>The following report displays results from AD computer account deactivation.</center></td></tr>
<tr class="noBorder"><td colspan=2></tr>
<tr><th>Action</th><th>Result</th></tr>
'  
#--[ Color presets for HTML ]--
$HexGray = "#dfdfdf"                                                   #--[ Grey default cell background ]--
$HexOrange = "#ff9900"                                                 #--[ Orange ]--
$HexYellow = "#ffd900"                                                 #--[ Yellow ]--
$HexBlack = "#000000"                                                  #--[ Black ]--
$HexGreen = "#006600"                                                  #--[ Green ]--
$HexRed = "#660000"                                                    #--[ Red ]--
                       
#--[ Data to Console display ]--------------------------------------
If ($Script:Console){ 
    write-host "Target System ="$TargetRecordObj.Hostname -ForegroundColor $Script:FgYellow 
    write-host "Today's Date ="$TargetRecordObj.Date -ForegroundColor $Script:FgCyan
    write-host "Technician ="$TargetRecordObj.Technician -ForegroundColor $Script:FgCyan
    write-host "Technician is authorized ="$TargetRecordObj.TechValid -ForegroundColor $Script:FgCyan
    write-host "AD Domain Name ="$TargetRecordObj.DomainName -ForegroundColor $Script:FgCyan          
    write-host "AD Status ="$TargetRecordObj.ADStatus -ForegroundColor $Script:FgCyan      
    write-host "Distinguished Name ="$TargetRecordObj.DistinguishedName -ForegroundColor $Script:FgCyan 
    write-host "Object Class ="$TargetRecordObj.ObjectClass -ForegroundColor $Script:FgCyan
    write-host "Object GUID ="$TargetRecordObj.ObjectGUID -ForegroundColor $Script:FgCyan
    write-host "SAM Account Name ="$TargetRecordObj.SamAccountName -ForegroundColor $Script:FgCyan
    write-host "Domain SID ="$TargetRecordObj.SID -ForegroundColor $Script:FgCyan
    write-host "AD Action ="$TargetRecordObj.AdAction -ForegroundColor $Script:FgCyan
    write-host "AD Action Result ="$TargetRecordObj.AdResult -ForegroundColor $Script:FgCyan
    write-host "Selected DNS Server ="$TargetRecordObj.DNSServer -ForegroundColor $Script:FgCyan         
    write-host "DNS Detection ="$TargetRecordObj.DNSDetect -ForegroundColor $Script:FgCyan         
    write-host "DNS PS Session ="$TargetRecordObj.DnsSession -ForegroundColor $Script:FgCyan
    write-host "DNS Action ="$TargetRecordObj.DnsAction -ForegroundColor $Script:FgCyan
    write-host "DNS Action Result ="$TargetRecordObj.DnsResult -ForegroundColor $Script:FgCyan
}

#--[ Data for email report ]--------------------------------------------
$Script:MessageBody += '<tr><td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>Target System</td>'
$Script:MessageBody += '<td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>' + $TargetRecordObj.Hostname + '</td></tr>'
$Script:MessageBody += '<tr><td bgcolor=' + $HexGray + '><font color=' + $HexBlack + ">Today's date</td>"
$Script:MessageBody += '<td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>' + $TargetRecordObj.Date + '</td></tr>'
$Script:MessageBody += '<tr><td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>Technician</td>'
$Script:MessageBody += '<td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>' + $TargetRecordObj.Technician + '</td></tr>'
$Script:MessageBody += '<tr><td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>Technician is authorized</td>'
$Script:MessageBody += '<td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>' + $TargetRecordObj.TechValid + '</td></tr>'
$Script:MessageBody += '<tr><td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>AD Domain Name</td>'
$Script:MessageBody += '<td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>' + $TargetRecordObj.DomainName + '</td></tr>'
$Script:MessageBody += '<tr><td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>AD Status</td>'
$Script:MessageBody += '<td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>' + $TargetRecordObj.ADStatus + '</td></tr>'
$Script:MessageBody += '<tr><td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>Distinguished Name</td>'
$Script:MessageBody += '<td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>' + $TargetRecordObj.DistinguishedName + '</td></tr>'
$Script:MessageBody += '<tr><td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>Object Class</td>'
$Script:MessageBody += '<td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>' + $TargetRecordObj.ObjectClass + '</td></tr>'
$Script:MessageBody += '<tr><td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>Object GUID</td>'
$Script:MessageBody += '<td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>' + $TargetRecordObj.ObjectGUID + '</td></tr>'
$Script:MessageBody += '<tr><td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>SAM Account Name</td>'
$Script:MessageBody += '<td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>' + $TargetRecordObj.SamAccountName + '</td></tr>'
$Script:MessageBody += '<tr><td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>Domain SID</td>'
$Script:MessageBody += '<td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>' + $TargetRecordObj.SID + '</td></tr>'
$Script:MessageBody += '<tr><td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>AD Action</td>'
$Script:MessageBody += '<td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>' + $TargetRecordObj.AdAction + '</td></tr>'
$Script:MessageBody += '<tr><td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>AD Result</td>'
$Script:MessageBody += '<td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>' + $TargetRecordObj.AdResult + '</td></tr>'
$Script:MessageBody += '<tr><td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>Selected DNS Server</td>'
$Script:MessageBody += '<td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>' + $TargetRecordObj.DNSServer + '</td></tr>'
$Script:MessageBody += '<tr><td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>DNS Record Detection</td>'
$Script:MessageBody += '<td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>' + $TargetRecordObj.DNSDetect + '</td></tr>'
$Script:MessageBody += '<tr><td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>DNS Server Session</td>'
$Script:MessageBody += '<td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>' + $TargetRecordObj.DnsSession + '</td></tr>'
$Script:MessageBody += '<tr><td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>DNS Record Action</td>'
$Script:MessageBody += '<td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>' + $TargetRecordObj.DnsAction + '</td></tr>'
$Script:MessageBody += '<tr><td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>DNS Record Result</td>'
$Script:MessageBody += '<td bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>' + $TargetRecordObj.DnsResult + '</td></tr>'
$Script:MessageBody += '<tr><td colspan=2 bgcolor=' + $HexGray + '><font color=' + $HexBlack + '>Target processing completed...</td></tr></table>'
If ($Script:Console -or $Script:Debug){
    $Script:MessageBody += '<br><br>Script Name : ' + $ScriptName   
    $Script:MessageBody += '<br>Script Executed On : ' + $TargetRecordObj.Date 
    $Script:MessageBody += '<br>Script Executed From : ' + $Env:ComputerName
    $Script:MessageBody += '<br>Script Executed By : ' + $TargetRecordObj.Technician
    $Script:MessageBody += '<br>Script Version : ' + $ScriptVer 
}    

SendMessage
If ($Script:Console){Write-Host "--- Completed ---`n" -ForegroundColor red }


<#==[ XML Configuration file example. Must reside in same folder as the script ]=======================
 
<!-- Settings & Configuration File -->
<Settings>
    <General>
        <ReportName>Computer Decommission Report</ReportName>
        <DebugTarget>test-computer</DebugTarget>
        <Domain>mydomain.com</Domain>
        <ValidGroup>TheGoodADGroup</ValidGroup>
    </General>
    <Email>
        <From>Computer-Decommission-Report@mydomain.com</From>
        <To>me@mydomain.com</To>
        <Debug>you@yourdomain.com</Debug>
        <Subject>Computer Decommission Report</Subject>
        <HTML>$true</HTML>
        <SmtpServer>10.10.50.5</SmtpServer>
    </Email>
    <Credentials>
        <UserName>mydomain\serviceaccount</UserName>
        <Password>76492d111IAegB2AHYAZQAxAGIATgBaADcAYwBtAHAAWQB6AHoAIAegB2AHYAZQQA9AIAegB2AHYAZQAxAGIATgBaADcAYwBtAHAAWHwAYwAzADQANgA0AGEAMAAwADQAZgBiAGMAYQBhAG6AHoAIAegB2AHYAZQQA9AIAegB2AHYAZ6AHoAIAegB2AHYAZQQA9AIAegB2AHYAZGUAZgBkAGYAZAA=</Password>
        <Key>kdhCh7HCvLOEyj2N0IObie8m+EhCh7HCvLO+Eyj2N0IObie8mE=</Key>
    </Credentials>
    <Bypass>
        <List>"PKIServer","Template1"</List> <!-- List of records to leave alone -->
    </Bypass>
</Settings>
#>