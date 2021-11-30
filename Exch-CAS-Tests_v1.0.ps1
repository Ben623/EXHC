########################################################################
# Created By  : Tzahi Kolber
# Version     : v1.0.0.0
# Last Update : 24/08/2021 11:25
########################################################################
# Disclaimer:
# The sample scripts are not supported under any Microsoft standard support program or service.
# The sample scripts are provided AS IS without warranty of any kind.
# Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. In no event shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.
########################################################################


# Export all data output to a log file #
Start-Transcript -path "C:\Scripts\ExchCasTest.Log"
  

#Add-PSSnapin *exch*
$exsrvs=(Get-mailboxserver).name
$FQDNNAME = (Get-WmiObject -Class Win32_ComputerSystem).domain

# allow TLS, TLS 1.1 and TLS 1.2 connections

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12


# Skip SSL certificate checks #

add-type @"
    using System.Net;
    using System.Security.Cryptography.X509Certificates;
    public class TrustAllCertsPolicy : ICertificatePolicy {
        public bool CheckValidationResult(
            ServicePoint srvPoint, X509Certificate certificate,
            WebRequest request, int certificateProblem) {
            return true;
        }
    }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy


########################### Start runnnig the tests ###########################

foreach ($srv in $exsrvs) {


#Get Vdir on each server #


#$vdir = "http"+":"+"//"+"$srv"+"."+"$FQDNNAME/"
$Mapivdir = "https"+":"+"//"+"$srv"+"."+"$FQDNNAME"+"/"+"mapi"+"/"+"Healthcheck"+"."+"htm"
#$MapivdirBackend = "https"+":"+"//"+"$srv"+"."+"$FQDNNAME"+":"+"444"+"/"+"mapi"+"/"+"Healthcheck"+"."+"htm"
$RPCvdir = "https"+":"+"//"+"$srv"+"."+"$FQDNNAME"+"/RPC/"+"Healthcheck"+"."+"htm"
$OABvdir = "https"+":"+"//"+"$srv"+"."+"$FQDNNAME"+"/OAB/"+"Healthcheck"+"."+"htm"
$EWSvdir = "https"+":"+"//"+"$srv"+"."+"$FQDNNAME"+"/EWS/"+"Healthcheck"+"."+"htm"
$Autodiscovervdir = "https"+":"+"//"+"$srv"+"."+"$FQDNNAME"+"/Autodiscover/"+"Healthcheck"+"."+"htm"
$OWAvdir = "https"+":"+"//"+"$srv"+"."+"$FQDNNAME"+"/OWA/"+"Healthcheck"+"."+"htm"
$ECPvdir = "https"+":"+"//"+"$srv"+"."+"$FQDNNAME"+"/ECP/"+"Healthcheck"+"."+"htm"
$EASvdir = "https"+":"+"//"+"$srv"+"."+"$FQDNNAME"+"/Microsoft-Server-ActiveSync/"+"Healthcheck"+"."+"htm"
$PowerShellvdir = "https"+":"+"//"+"$srv"+"."+"$FQDNNAME"+"/PowerShell/"+"Healthcheck"+"."+"htm"



# Run Ping Test as a Master test #
Write-Host -ForegroundColor Cyan "Run Ping Test for $srv"

if (Test-Connection $srv -Count 2 -ErrorAction SilentlyContinue) {

     write-host  -ForegroundColor Green "Computer $srv responded to ping test"

# Continue to all other tests #       
        
       

Write-Host -ForegroundColor Cyan "Testing $srv Virtual Direcotries State"
# Run Web Test for MAPI Vdir #

try {
$Mapiresponce = Invoke-WebRequest -Uri $Mapivdir -UseBasicParsing
$statcode = $Mapiresponce.statuscode
#Write-host -ForegroundColor Cyan $statcode
$statdes = $Mapiresponce.StatusDescription
#Write-host -ForegroundColor Cyan $statdes
If ($statcode -eq "200" -and $statdes -eq "OK") {
Write-Host -ForegroundColor Green "WebSite - $MapiVdir is up (Return code: $statcode - $statdes)"
}
else {
Write-Host -ForegroundColor Red "WebSite - $MapiVdir is DOWN (Return code: $statcode - $statdes)"
}
} catch {
Write-Host -ForegroundColor Red "WebSite - $MapiVdir is not accessible. Please check"
}

# Run Web Test for MapivdirBackend Vdir
#try {
#$MapivdirBackendresponce = Invoke-WebRequest -Uri $MapivdirBackend
#$statcode = $MapivdirBackendresponce.statuscode
#Write-host -ForegroundColor Cyan $statcode
#$statdes = $MapivdirBackendresponce.StatusDescription
#Write-host -ForegroundColor Cyan $statdes
#If ($statcode -eq "200" -and $statdes -eq "OK") {
#Write-Host -ForegroundColor Green "WebSite - $MapivdirBackend is up (Return code: $statcode - $statdes)"
#}
#else {
#Write-Host -ForegroundColor Red "WebSite - $MapivdirBackend is DOWN (Return code: $statcode - $statdes)"
#}
#} catch {
#Write-Host -ForegroundColor Red "WebSite - $MapivdirBackendVdir is not accessible. PlMapivdirBackende check"
#}


# Run Web Test for RPC Vdir
try {
$RPCresponce = Invoke-WebRequest -Uri $RPCvdir -UseBasicParsing
$statcode = $RPCresponce.statuscode
#Write-host -ForegroundColor Cyan $statcode
$statdes = $RPCresponce.StatusDescription
#Write-host -ForegroundColor Cyan $statdes
If ($statcode -eq "200" -and $statdes -eq "OK") {
Write-Host -ForegroundColor Green "WebSite - $RPCVdir is up (Return code: $statcode - $statdes)"
}
else {
Write-Host -ForegroundColor Red "WebSite - $RPCVdir is DOWN (Return code: $statcode - $statdes)"
}
} catch {
Write-Host -ForegroundColor Red "WebSite - $RPCVdir is not accessible. Please check"
}

# Run Web Test for OAB Vdir
try {
$OABresponce = Invoke-WebRequest -Uri $OABvdir -UseBasicParsing
$statcode = $OABresponce.statuscode
#Write-host -ForegroundColor Cyan $statcode
$statdes = $OABresponce.StatusDescription
#Write-host -ForegroundColor Cyan $statdes
If ($statcode -eq "200" -and $statdes -eq "OK") {
Write-Host -ForegroundColor Green "WebSite - $OABVdir is up (Return code: $statcode - $statdes)"
}
else {
Write-Host -ForegroundColor Red "WebSite - $OABVdir is DOWN (Return code: $statcode - $statdes)"
}
} catch {
Write-Host -ForegroundColor Red "WebSite - $OABVdir is not accessible. Please check"
}

# Run Web Test for EWS Vdir
try {
$EWSresponce = Invoke-WebRequest -Uri $EWSvdir -UseBasicParsing
$statcode = $EWSresponce.statuscode
#Write-host -ForegroundColor Cyan $statcode
$statdes = $EWSresponce.StatusDescription
#Write-host -ForegroundColor Cyan $statdes
If ($statcode -eq "200" -and $statdes -eq "OK") {
Write-Host -ForegroundColor Green "WebSite - $EWSVdir is up (Return code: $statcode - $statdes)"
}
else {
Write-Host -ForegroundColor Red "WebSite - $EWSVdir is DOWN (Return code: $statcode - $statdes)"
}
} catch {
Write-Host -ForegroundColor Red "WebSite - $EWSVdir is not accessible. Please check"
}


# Run Web Test for Autodiscover Vdir
try {
$Autodiscoverresponce = Invoke-WebRequest -Uri $Autodiscovervdir -UseBasicParsing
$statcode = $Autodiscoverresponce.statuscode
#Write-host -ForegroundColor Cyan $statcode
$statdes = $Autodiscoverresponce.StatusDescription
#Write-host -ForegroundColor Cyan $statdes
If ($statcode -eq "200" -and $statdes -eq "OK") {
Write-Host -ForegroundColor Green "WebSite - $AutodiscoverVdir is up (Return code: $statcode - $statdes)"
}
else {
Write-Host -ForegroundColor Red "WebSite - $AutodiscoverVdir is DOWN (Return code: $statcode - $statdes)"
}
} catch {
Write-Host -ForegroundColor Red "WebSite - $AutodiscoverVdir is not accessible. Please check"
}

# Run Web Test for OWA Vdir
try {
$OWAresponce = Invoke-WebRequest -Uri $OWAvdir -UseBasicParsing
$statcode = $OWAresponce.statuscode
#Write-host -ForegroundColor Cyan $statcode
$statdes = $OWAresponce.StatusDescription
#Write-host -ForegroundColor Cyan $statdes
If ($statcode -eq "200" -and $statdes -eq "OK") {
Write-Host -ForegroundColor Green "WebSite - $OWAVdir is up (Return code: $statcode - $statdes)"
}
else {
Write-Host -ForegroundColor Red "WebSite - $OWAVdir is DOWN (Return code: $statcode - $statdes)"
}
} catch {
Write-Host -ForegroundColor Red "WebSite - $OWAVdir is not accessible. Please check"
}


# Run Web Test for ECP Vdir
try {
$ECPresponce = Invoke-WebRequest -Uri $ECPvdir -UseBasicParsing
$statcode = $ECPresponce.statuscode
#Write-host -ForegroundColor Cyan $statcode
$statdes = $ECPresponce.StatusDescription
#Write-host -ForegroundColor Cyan $statdes
If ($statcode -eq "200" -and $statdes -eq "OK") {
Write-Host -ForegroundColor Green "WebSite - $ECPVdir is up (Return code: $statcode - $statdes)"
}
else {
Write-Host -ForegroundColor Red "WebSite - $ECPVdir is DOWN (Return code: $statcode - $statdes)"
}
} catch {
Write-Host -ForegroundColor Red "WebSite - $ECPVdir is not accessible. Please check"
}

# Run Web Test for EAS Vdir
try {
$EASresponce = Invoke-WebRequest -Uri $EASvdir -UseBasicParsing
$statcode = $EASresponce.statuscode
#Write-host -ForegroundColor Cyan $statcode
$statdes = $EASresponce.StatusDescription
#Write-host -ForegroundColor Cyan $statdes
If ($statcode -eq "200" -and $statdes -eq "OK") {
Write-Host -ForegroundColor Green "WebSite - $EASVdir is up (Return code: $statcode - $statdes)"
}
else {
Write-Host -ForegroundColor Red "WebSite - $EASVdir is DOWN (Return code: $statcode - $statdes)"
}
} catch {
Write-Host -ForegroundColor Red "WebSite - $EASVdir is not accessible. Please check"
}

# Run Web Test for Powershell Vdir
try {
$Powershellresponce = Invoke-WebRequest -Uri $Powershellvdir -UseBasicParsing
$statcode = $Powershellresponce.statuscode
#Write-host -ForegroundColor Cyan $statcode
$statdes = $Powershellresponce.StatusDescription
#Write-host -ForegroundColor Cyan $statdes
If ($statcode -eq "200" -and $statdes -eq "OK") {
Write-Host -ForegroundColor Green "WebSite - $PowershellVdir is up (Return code: $statcode - $statdes)"
}
else {
Write-Host -ForegroundColor Red "WebSite - $PowershellVdir is DOWN (Return code: $statcode - $statdes)"
}
} catch {
Write-Host -ForegroundColor Red "WebSite - $PowershellVdir is not accessible. PlPowershelle check"
}

#Write-Host -ForegroundColor Cyan "Testing $srv Virtual Direcotries State - ENDED"


# Run Server Commponent State Test #

Write-Host -ForegroundColor Cyan "Testing $srv Commponents State"

$compstate = Get-ServerComponentState $srv | Select-Object Component,State
foreach ($com in $compstate)
{
$comsrv = $com.Component
$comstat = $com.State
If ($comstat -eq "Inactive") {

Write-Host -ForegroundColor Red "$comsrv on $srv is $comstat"
    }
        }

#Write-Host -ForegroundColor Cyan "Testing $srv Commponents State - ENDED"


# Run Client Access Ports tests #
Write-Host -ForegroundColor Cyan "Testing $srv Client Access Ports"

$Testports = 80,443,444
$Testports | ForEach-Object {
$Testports = $_; if (Test-NetConnection -ComputerName $srv -Port $Testports -InformationLevel Quiet -WarningAction SilentlyContinue) 
{Write-Host -ForegroundColor Green "Port $Testports is open" } 
else {Write-Host -ForegroundColor Red "Port $Testports is closed"}
 }  
#Write-Host -ForegroundColor Cyan "Testing $srv Client Access Ports - ENDED"


# Run Certificates expiration tests #

Write-Host -ForegroundColor Cyan "Testing $srv Certificates expiration"

$certstate = Get-ExchangeCertificate -Server $srv | Select-Object subject,NotAfter
foreach ($cert in $certstate)
{
$certdate = ($cert.NotAfter).ToString("dddd, yyyy-MM-dd")
$nowdate = Get-date -Format "dddd, yyyy-MM-dd"
#$certdt = ($cert.NotAfter) 
$daysremain = (New-TimeSpan -Start $nowdate -End $certdate).Days
If ($daysremain -le "40") {
$certname = $cert.Subject
Write-Host -ForegroundColor Red "$certname certificate is about to expire in on $srv in $daysremain days"
    }
        }

#Write-Host -ForegroundColor Cyan "Certificates expiration tests - ENDED"


Write-Host -ForegroundColor Cyan "Certificates existence tests"

# Run Certificates existence tests #
$session = New-PsSession –ComputerName $srv
#Write-Host $session.ComputerName
#Enter-PSSession $session
#Import-Module WebAdministration
Invoke-Command -Session $session -ScriptBlock {Import-Module WebAdministration ; $t444temp = Get-ChildItem -Path IIS:SSLBindings | Where-Object {$_.Port -eq "444"} | Select-Object Thumbprint}
$t444 = Invoke-Command -Session $session -ScriptBlock {$t444temp.Thumbprint}
#Write-Host $t444
$t444all = Invoke-Command -Session $session -ScriptBlock {Import-Module WebAdministration ; Get-ChildItem -Path CERT:LocalMachine/My}
#Write-Host $t444all
If (($t444all.Thumbprint) -notcontains $t444) {Write-Host -ForegroundColor red "There is no Certificate binded to Exchange Back End site port 444"}
Else {Write-Host -ForegroundColor Green "Certificate is binded to Exchange Back End site port 444"}

Invoke-Command -Session $session -ScriptBlock {Import-Module WebAdministration ; $t443temp = Get-ChildItem -Path IIS:SSLBindings | Where-Object {$_.Port -eq "443"} | Select-Object Thumbprint}
$t443 = Invoke-Command -Session $session -ScriptBlock {$t443temp.Thumbprint | Select-Object -First 1}
#Write-Host $t443
$t443all = Invoke-Command -Session $session -ScriptBlock {Import-Module WebAdministration ; Get-ChildItem -Path CERT:LocalMachine/My}
#Write-Host $t443all.Thumbprint
If (($t443all.Thumbprint) -notcontains $t443) {Write-Host -ForegroundColor red "There is no Certificate binded to Exchange Default web site port 443"}
Else {Write-Host -ForegroundColor Green "Certificate is binded to Exchange Default web site port 443"}


#Write-Host -ForegroundColor Cyan "Certificates existence tests - ENDED"


Write-Host -ForegroundColor Cyan "Run Application Pools state tests"

# Run Application Pools state #
#$session = New-PsSession –ComputerName $srv
$session = New-PsSession –ComputerName $srv
Invoke-Command -Session $session -ScriptBlock {$apitem=@(Get-WebAppPoolState | Select-Object value,ItemXPath)}
$apstate = Invoke-Command -Session $session -ScriptBlock {$apitem}
ForEach ($aps in $apstate) {
$te=($aps).itemxpath
$ten = $te.split("@") | Select-Object -Last 1
$tenew = ($ten.split("''") )[1]
If (($aps.value) -eq "Stopped") {Write-Host -ForegroundColor red $tenew "is Stopped"}
Else {Write-Host -ForegroundColor Green $tenew "is running"}
}


#Write-Host -ForegroundColor Cyan "Application Pools state # - ENDED"


# Run Mapi Connectivity tests #

Write-Host -ForegroundColor Cyan "Testing Mapi Connectivity on $srv"

$Mapistate = Test-MAPIConnectivity -Server $srv -IncludePassive:$true -MonitoringContext:$true
$mapisresult = $Mapistate.Result.value
$mapiserr = $Mapistate.Error
$mapisevent = $Mapistate.Events
If (($mapisresult -eq "Success") -and ($mapiserr -eq "$null")) {Write-Host -ForegroundColor Green "Mapi tests completed successfully on $SRV"}
    Else {
    Write-Host -ForegroundColor Red "$srv is having a Mapi test issues - Please check"
    Write-Host -ForegroundColor Red $mapisevent}

#Write-Host -ForegroundColor Cyan "Mapi Connectivity tests - ENDED"


 } else {Write-host -ForegroundColor Red "The computer $srv could not be contacted"}

}

Stop-Transcript