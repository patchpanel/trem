#powershell.exe -executionpolicy bypass -file  c:\trem\bin/SendMailIndividual.ps1 "c:\trem\in/201603 - BDG_TimeReport_V2.xlsx" "jl186034@teradata.com" "jl186034@teradata.com" "localhost" "201604" "in-script" "c:\trem\log" "Practice"
#Badge Reports Manila QLID: bm230103
param (
    [string]$argsExceBadgeReport = $(throw "-Excel Badge Report file is required."),
    [string]$argsFromEmailAddress = $(throw "-FromEmailAddress is required."),
	[string]$argsToEmailAddress = $(throw "-FromEmailAddress is required."),
    [string]$argsSMTPServer = $(throw "-EmailSMPTServer is required."),
	[string]$argsBatchID = $(throw "-BatchID is required."),
	[string]$argsEmailBody = $(throw "-Email Body is required."),
	[string]$argsLogDir = $(throw "-Log Directory is required."),
	[string]$argsMngrRptTag = $(throw "-Manager tag is required."),
	[String]$argsTempDir = $(throw "-Temp Directory is required.")
)


Function Zip-File {
	param(
	[String]$inputFile,
	[String]$outputFile
)
	& "C:\Program Files\7-Zip\7z.exe" u -mx9 -t7z -m0=lzma2 ("$(split-path $inputFile)\$outputFile") $argsExceBadgeReport
	if ($LASTEXITCODE -eq 0) {
		#Remove-Item $argsExceBadgeReport -Force -Recurse -ErrorAction SilentlyContinue
		Write-Host -ForegroundColor DarkMagenta "[$(Get-Date)] Successful Zip"
	} else {
		Exit 99
	}
}

Set-Variable -Name emailBody -Visibility Private
Set-Variable -Name fileName -Visibility Private
Set-Variable -Name emailAddress -Visibility Private
Set-Variable -Name emailSubject -Visibility Private
Set-Variable -Name emailAttachments -Visibility Private
Set-Variable -Name flgEAexists -Visibility Private
Set-Variable -Name resultsFile -Visibility Private
Set-Variable -Name logFile -Visibility Private
Set-Variable -Name errorActionPreference -Visibility Private
Set-Variable -Name message -Visibility Private
Set-Variable -Name nCtr -Visibility Private
Set-Variable -Name mCtr -Visibility Private

#Get the timing
$elapsed = [System.Diagnostics.Stopwatch]::StartNew()
$started = Get-Date

Write-Host -ForegroundColor Green "=================================================================="
Write-Host -ForegroundColor DarkMagenta "[$(Get-Date)] Sending complete Badge report"
Write-Host -ForegroundColor Green "=================================================================="

try
{
	$ProcessedDate = [datetime]::ParseExact($argsBatchID,'yyyyMM',[System.Globalization.CultureInfo]::CurrentCulture)
	$ProcessedDate = $ProcessedDate.ToString("MMMM yyyy")
}
catch [System.Exception]
{
    Write-Host  -ForegroundColor Red "You have entered an Invalid Batch ID. Date format should be YYYYMM"
	Exit 99
}
#Check if reporting period is closed.
$lockPath = "$env:APPDATA\trem"
if (($(Test-Path "$lockPath\tremind.$argsBatchID.lck") -eq 1) -and ($(Test-Path "$lockPath\tremgrp.$argsBatchID.lck") -eq 1)) {
	Write-Host  -ForegroundColor Red "[$(Get-Date)] Reporting Period $argsBatchID is already closed. Please Enter a new period to proceed."
	Exit 69;
}

#Don't continue if there unsent emails from previous script
$indLogFile = "smtp.$argsBatchID.log"
$grpLogFile = "smtp.$argsBatchID.$argsMngrRptTag.log"
if (($(Test-Path "$argsLogDir\$grpLogFile") -eq 1) -or ($(Test-Path "$argsLogDir\$indLogFile") -eq 1)) {
			Write-Host -ForegroundColor DarkRed "[$(Get-Date)] Cannot continue. Individual or Practice Reports have unsent attachments."
			Exit 98
}

#Build Domain name
$aSMTPServer = $argsSMTPServer.Split(".")
$tld = $aSMTPServer[$aSMTPServer.GetUpperBound(0)]
$domain = $aSMTPServer[$aSMTPServer.GetUpperBound(0)-1] + "." + $tld
#Get and Save credentials before sending emails
$pw = $null
$cred = $null
if ($(Test-Path "$argsTempDir\$argsBatchID.pw") -eq 0)
{
	(Get-Credential).password | ConvertFrom-SecureString > "$argsTempDir\$argsBatchID.pw"
    Write-Host -ForegroundColor Cyan "$argsTempDir\$argsBatchID.pw created"
}
$pw = Get-Content "$argsTempDir\$argsBatchID.pw" | ConvertTo-SecureString
$cred = New-Object System.Management.Automation.PSCredential "MailUser", $pw

#Check the Badge report file
$flgBadgeReptExists = Test-Path $argsExceBadgeReport
if ($flgBadgeReptExists -eq 0)
{
    Write-Host -ForegroundColor Red "$argsExceBadgeReport not found!"
    Exit 99
}

#20160520: Compress file. 
#It might be easier to send it via SMTP
$zipFile = ([io.fileinfo]"$argsExceBadgeReport").basename
Zip-File $argsExceBadgeReport $zipFile

#+-------------------------------------------------------+'
#|                 SEND INDIVIDUAL REPORT                 |'
#+-------------------------------------------------------+'
$errorActionPreference = "Stop" 
$nCtr = 0
$mCtr = 0
$logFile = "smtp.Single.$argsBatchID.log"
#$emailBody = $argsEmailBody
#Hardcode values for the mean-time
$emailBody = "Hi,`r`n`r`nAttached is the Badge Report for the aforementioned month.`r`n`r`nManila Badge Report`r`n`r`n`r`n`r`n***System generated email. Please do not reply.***"
	#$fileName = $argsExceBadgeReport
	$fileName = "$(split-path $argsExceBadgeReport)\$zipFile.7z"
	Write-Host $fileName
	$emailAddress = $argsToEmailAddress
	$emailSubject = "Badge Report for $ProcessedDate"
	$emailAttachments = $fileName
	$flgEAexists = Test-Path $emailAttachments
	if ($flgEAexists -eq 1) {
		$logFile = "$argsBatchID.smtp.Individual.log" 
		try
		{             
			Write-host -ForegroundColor Cyan "[$(Get-Date)] Sending to $emailAddress" 
			if (($argsSMTPServer.ToLower() -eq "localhost") -or ($argsSMTPServer.ToLower() -eq "127.0.0.1")) {
				send-mailmessage -ErrorAction Stop `
				-from $argsFromEmailAddress `
				-to $emailAddress `
				-subject $emailSubject `
				-smtpServer $argsSMTPServer `
				-body $emailBody `
				-Attachments $emailAttachments `
				-DeliveryNotificationOption OnFailure `
				-Priority High `
				-Port 25 `
			} else {
				send-mailmessage -ErrorAction Stop `
				-Credential $cred `
				-from $argsFromEmailAddress `
				-to $emailAddress `
				-subject $emailSubject `
				-smtpServer $argsSMTPServer `
				-body $emailBody `
				-Attachments $emailAttachments `
				-DeliveryNotificationOption OnFailure `
				-Priority High `
				-Port 25
			}
			$mCtr++
		}                        
		catch [System.Exception]
		{		
			#Record Failed sending
			$fileName | Out-File "$argsLogDir\$logFile" -Encoding unicode -Append
			#Log exception details
			$_.Exception.GetType().FullName + "`r`n" + `
			$_.Exception.Message + "`r`n" + `
			$_.Exception.Stacktrace `
			| Add-Content "$argsLogDir\$logFile" -Encoding unicode
			#echo $resources + " " + $_.Exception | format-list -force >> "$argsLogDir\$logFile" 
			Write-host  -ForegroundColor Red "[$(Get-Date)] Oooops. There was an error in sending $emailAttachments. Please check $logFile"
			
            $nCtr++
		}           
	}
	else {
		Write-host  -ForegroundColor Red "$emailAttachments =>Attachment not Found!"
		$nCtr++
}
Write-Host -ForegroundColor Green "=================================================================="
Write-Host -ForegroundColor DarkMagenta "Total Sent: $mCtr"
Write-Host -ForegroundColor DarkMagenta "Total UnSent: $nCtr"
Write-host "Started at $started"
Write-host "Ended at $(get-date)"
Write-host "Total Elapsed Time: $($elapsed.Elapsed.ToString())"
Write-Host -ForegroundColor Green "=================================================================="

if ($nCtr -gt 0) {
	Exit 99
}