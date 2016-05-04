#powershell.exe -executionpolicy bypass -file  c:\trem\bin/SendMailAll.ps1 "c:\trem\in/ResourceList.txt" "c:\trem\in/ManagersList.txt" "C:\trem\out" "jl186034@teradata.com" "localhost" "201604" "Practice" "in-script" "c:\trem\log"
#Badge Reports Manila QLID: bm230103
param (
    [String]$argsResourceList = $(throw "-Resource List is required."),
    [String]$argsManagerList = $(throw "-Manager List is required."), 
    [String]$argsExcelOutputDir = $(throw "-Excel output Directory is required."),
    [String]$argsFromEmailAddress = $(throw "-FromEmailAddress is required."),
    [String]$argsSMTPServer = $(throw "-EmailSMPTServer is required."),
    [String]$argsBatchID = $(throw "-BatchID is required."),
    [String]$argsMngrRptTag = $(throw "-Manager Report Tag is required."),
    [String]$argsEmailBody = $(throw "-Email Body is required."),
    [String]$argsLogDir = $(throw "-Log Directory is required."),
	[String]$argsTempDir = $(throw "-Temp Directory is required.")
)

Set-Variable -Name emailBody -Visibility Private
Set-Variable -Name lineResources -Visibility Private
Set-Variable -Name aResources -Visibility Private
Set-Variable -Name qlid -Visibility Private
Set-Variable -Name resourceName -Visibility Private
Set-Variable -Name fileName -Visibility Private
Set-Variable -Name emailAddress -Visibility Private
Set-Variable -Name emailSubject -Visibility Private
Set-Variable -Name emailAttachments -Visibility Private
Set-Variable -Name flgEAexists -Visibility Private
Set-Variable -Name unsentFile -Visibility Private
Set-Variable -Name logFile -Visibility Private
Set-Variable -Name message -Visibility Private
Set-Variable -Name nCtr -Visibility Public
Set-Variable -Name mCtr -Visibility Public
Set-Variable -Name repro -Visibility Private
Set-Variable -Name isoDate -Value Private

#Sent/Unsent counters
$mCtr = 0
$nCtr = 0

function Send-Mail-Group {
    param(
    [String]$resources,
	[System.Management.Automation.PSCredential]$credentials
)
    $aResources = $resources.Split("|")
    $qlid = $aResources[0].ToLower().Trim()
    $resourceName = $aResources[1].Trim()
    $fileName = $argsBatchID + ' ' + $argsMngrRptTag + ' ' + $resourceName
    $emailAddress = "$qlid@$domain"
    $emailSubject = "$argsMngrRptTag Badge Report for $ProcessedDate"
    $emailAttachments = $argsExcelOutputDir + '\' + $fileName + '.xlsx'
    $flgEAexists = Test-Path $emailAttachments
   
    if ($flgEAexists -eq 1) {
			Write-host  -ForegroundColor Cyan "[$(Get-Date)] Sending to $resourceName."
            try {
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
					-Credential $credentials `
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
            $Global:mCtr++
            }  catch {
                #Record Failed sending
                $resources | Out-File "$argsLogDir\$unsentFile" -Encoding unicode -Append
                #Log exception details
                $_.Exception.GetType().FullName + "`r`n" + `
                $_.Exception.Message + "`r`n" + `
                $_.Exception.Stacktrace `
                | Add-Content "$argsLogDir\$logFile" -Encoding unicode
                #echo $resources + " " + $_.Exception | format-list -force >> "$argsLogDir\$logFile" 
                Write-host  -ForegroundColor Red "[$(Get-Date)] Oooops. There was an error in sending $emailAttachments. Please check $logFile"
                $Global:nCtr++
            }     
    } else {
        #Write-host  -ForegroundColor Red "[$(Get-Date)] $emailAttachments was not found!"
        #$Global:nCtr++
    }
}

function Send-Mail-Individual {
param(
    [String]$resources,
	[System.Management.Automation.PSCredential]$credentials
)
    #Build send-mailmessage params
    $aResources = $resources.Split("|")
    $qlid = $aResources[0].ToLower().Trim()
    $resourceName = $aResources[1].Trim()
    $fileName = $argsBatchID + " " + $resourceName
    $emailAddress = "$qlid@$domain"
    $emailSubject = "Badge Report for $ProcessedDate"
    $emailAttachments = $argsExcelOutputDir + '\' + $fileName + '.xlsx'
    $flgEAexists = Test-Path $emailAttachments
    #If attachment exists, go send it
    if ($flgEAexists -eq 1) {
			Write-host  -ForegroundColor Cyan "[$(Get-Date)] Sending to $resourceName."
            try {
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
					-Credential $credentials `
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
            $Global:mCtr++
            }  catch {
                #Record Failed sending
                $resources | Out-File "$argsLogDir\$unsentFile" -Encoding unicode -Append
                #Log exception details
                $_.Exception.GetType().FullName + "`r`n" + `
                $_.Exception.Message + "`r`n" + `
                $_.Exception.Stacktrace `
                | Add-Content "$argsLogDir\$logFile" -Encoding unicode
                #echo $resources + " " + $_.Exception | format-list -force >> "$argsLogDir\$logFile" 
                Write-host  -ForegroundColor Red "[$(Get-Date)] Oooops. There was an error in sending $emailAttachments. Please check $logFile"
                $Global:nCtr++
            }        
    } else {
        #Write-host  -ForegroundColor Red "[$(Get-Date)] $emailAttachments was not found!"
        #$Global:nCtr++
    }
}

#+-------------------------------------------------------+'
#|                 Main Program                          |'
#+-------------------------------------------------------+'

#Get the timing
$elapsed = [System.Diagnostics.Stopwatch]::StartNew()
$started = Get-Date

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

#Check if input files exist
#$flgResListExists = Test-Path $argsResourceList
if ($(Test-Path $argsResourceList) -eq 0)
{
    Write-Host -ForegroundColor Red "$argsResourceList not found!"
    Exit 99
}
#$flgMngrListExists = Test-Path $argsManagerList
if ($(Test-Path $argsManagerList) -eq 0)
{
    Write-Host -ForegroundColor Red "$argsManagerList not found!"
    Exit 99
}

if ($(Test-Path $argsExcelOutputDir) -eq 0)
{
    Write-Host -ForegroundColor Red "$argsExcelOutputDir not found!"
    Exit 99
}

if ($(Test-Path $argsLogDir) -eq 0)
{
    Write-Host -ForegroundColor Red "$argsLogDir not found!"
    Exit 99
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

#+-------------------------------------------------------+'
#|                 SEND INDIVIDUAL REPORTS               |'
#+-------------------------------------------------------+'

#Initialize Re-process flag
$repro = 0
#Remove logs and result file first
$logFile = "$argsBatchID.smtp.log"
$grpLogFile = "$argsBatchID.smtp.$argsMngrRptTag.log"
$unsentFile = "$argsBatchID.unsent.log"
$isoDate = Get-Date -format yyyyMMddHHmmss

if ($(Test-Path "$argsLogDir\$logFile") -eq 1)
{
    $repro = 1
    #Delete log file since sending email may generate another one
    #Backup the previous unsent file since sending email may produce
    #another one. We will be using this as reference not the original
    #resource list anymore. Why send them all again.
    Remove-Item "$argsLogDir\$logFile" -Force
    Copy-Item "$argsLogDir\$unsentFile" "$argsLogDir\$unsentFile.$isoDate" -Force
    Remove-Item "$argsLogDir\$unsentFile" -Force
} 

Write-Host -ForegroundColor Green "=================================================================="
Write-Host -ForegroundColor DarkMagenta "[$(Get-Date)] Sending Individual reports..."  
Write-Host -ForegroundColor Green "=================================================================="

#Check for complete processing.
#Skip if all are ok. Proceed to Per practice report
if ($(Test-Path "$lockPath\tremind.$argsBatchID.lck") -eq 0) {
    #$emailBody = $argsEmailBody
    #Hardcode values for the mean-time
    $emailBody = "Hi,`r`n`r`nAttached is your Badge report for the aforementioned month.`r`n`r`nManila Badge Report`r`n`r`n`r`n`r`n***System generated email. Please do not reply.***"

    if ($repro -eq 1) {
        #Use the backup file since resending emails will produce the original filename
        foreach ($line in Get-Content "$argsLogDir\$unsentFile.$isoDate") {
             Send-Mail-Individual -resources $line -credentials $cred
        }
        #Create lock file if completed. Transferred to Java GUI
        #if (($(Test-Path "$argsLogDir\$logFile") -eq 0) -and ($(Test-Path "$argsLogDir\$unsentFile") -eq 0) -and ($Global:nCtr -eq 0)) {
        #    $isoDate | Out-File "$lockPath\tremind.$argsBatchID.lck" -Encoding unicode -Append
        #}
    } else {
        #Regular processing should go here
        #Make sure no repro related files are there
		if ($(Test-Path "$argsLogDir\$grpLogFile") -eq 1) {
			Write-Host -ForegroundColor DarkRed "[$(Get-Date)] --Skipping--. Practice Reports have unsent attachments."  
        } else {
			Remove-Item "$argsLogDir\$logFile*" -Force
			Remove-Item "$argsLogDir\$unsentFile*" -Force
			foreach ($line in Get-Content $argsResourceList) {
				 Send-Mail-Individual -resources $line -credentials $cred
			}
		}
        #Create lock file if completed. Transferred to Java GUI
        #if (($(Test-Path "$argsLogDir\$logFile") -eq 0) -and ($(Test-Path "$argsLogDir\$unsentFile") -eq 0) -and ($Global:nCtr -eq 0)) {
        #    $isoDate | Out-File "$lockPath\tremind.$argsBatchID.lck" -Encoding unicode -Append
        #}
    }
} else {
    Write-Host "[$(Get-Date)] --Skipped. Already completed--"  
}

#+-------------------------------------------------------+'
#|                 SEND PRACTICE REPORTS                 |'
#+-------------------------------------------------------+'

Clear-Variable emailBody
Clear-Variable lineResources
Clear-Variable aResources
Clear-Variable qlid
Clear-Variable resourceName
Clear-Variable fileName
Clear-Variable emailAddress
Clear-Variable emailSubject
Clear-Variable emailAttachments
Clear-Variable flgEAexists
Clear-Variable unsentFile
Clear-Variable logFile
Clear-Variable message
Clear-Variable repro
Clear-Variable isoDate

#Initialize Re-process flag
$repro = 0
#Remove logs and result file first
$logFile = "$argsBatchID.smtp.$argsMngrRptTag.log"
$indLogFile = "$argsBatchID.smtp.log"
$unsentFile = "$argsBatchID.unsent.$argsMngrRptTag.log"
$isoDate = Get-Date -format yyyyMMddHHmmss

if ($(Test-Path "$argsLogDir\$logFile") -eq 1)
{
    $repro = 1
    #Delete log file since sending email may generate another one
    #Backup the previous unsent file since sending email may produce
    #another one. We will be using this as reference not the original
    #resource list anymore. Why send them all again.
    Remove-Item "$argsLogDir\$logFile" -Force
    Copy-Item "$argsLogDir\$unsentFile" "$argsLogDir\$unsentFile.$isoDate" -Force
    Remove-Item "$argsLogDir\$unsentFile" -Force
}
              
Write-Host -ForegroundColor Green "=================================================================="
Write-Host -ForegroundColor DarkMagenta "[$(Get-Date)] Sending Per Practice reports"  
Write-Host -ForegroundColor Green "=================================================================="

#Check for complete processing.
#Skip if all are ok. Proceed to Per practice report
if ($(Test-Path "tremgrp.$argsBatchID.lck") -eq 0) {
    $emailBody = $argsEmailBody
    #Hardcode values for the mean-time
    $emailBody = "Hi,`r`n`r`nAttached is the Per Practice Badge report for the aforementioned month.`r`n`r`nManila Badge Report`r`n`r`n`r`n`r`n***System generated email. Please do not reply.***"

    if ($repro -eq 1) {
        #Use the backup file since resending emails will produce the original filename
        foreach ($line in Get-Content "$argsLogDir\$unsentFile.$isoDate") {
            Send-Mail-Group -resources $line -credentials $cred
        }
        #Create lock file if completed. Transferred to Java GUI
        #if (($(Test-Path "$argsLogDir\$logFile") -eq 0) -and ($(Test-Path "$argsLogDir\$unsentFile") -eq 0) -and ($Global:nCtr -eq 0)) {
        #    $isoDate | Out-File "$lockPath\tremgrp.$argsBatchID.lck" -Encoding unicode -Append
        #}
    } else {
        #Regular processing should go here
        #Make sure no repro related files are there
		#If individual reports have errors, don't process yet
		if ($(Test-Path "$argsLogDir\$indLogFile") -eq 1) {
			Write-Host -ForegroundColor DarkRed "[$(Get-Date)] Cannot continue. Individual Reports have unsent attachments."  
			Exit 98
		} else {
			Remove-Item "$argsLogDir\$logFile*" -Force
			Remove-Item "$argsLogDir\$unsentFile*" -Force
			foreach ($line in Get-Content $argsManagerList) {
				Send-Mail-Group -resources $line -credentials $cred
			}
		}
        #Create lock file if completed. Transferred to Java GUI
        #if (($(Test-Path "$argsLogDir\$logFile") -eq 0) -and ($(Test-Path "$argsLogDir\$unsentFile") -eq 0) -and ($Global:nCtr -eq 0)) {
        #    $isoDate | Out-File "$lockPath\tremgrp.$argsBatchID.lck" -Encoding unicode -Append
        #}
    }
} else {
    Write-Host "[$(Get-Date)] --Skipped. Already completed--"  
}

#+-------------------------------------------------------+'
#|                 Report Status                         |'
#+-------------------------------------------------------+'

Write-Host -ForegroundColor Green "=================================================================="
Write-Host -ForegroundColor DarkMagenta "Total Sent: $mCtr"
Write-Host -ForegroundColor DarkMagenta "Total UnSent: $nCtr"
Write-host "Started at $started"
Write-host "Ended at $(get-date)"
Write-host "Total Elapsed Time: $($elapsed.Elapsed.ToString())"
Write-Host -ForegroundColor Green "=================================================================="