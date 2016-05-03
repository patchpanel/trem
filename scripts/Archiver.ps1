# Powershell script - move old files to an archive location. 
# Writes log files to $logPath 
# Ver 0.6 
#powershell -ExecutionPolicy ByPass -File .\archiver.ps1 "c:\brem\out" "c:\brem\archive" "c:\brem\log" "OUT" 1
#powershell -ExecutionPolicy ByPass -File .\archiver.ps1 "c:\brem\in" "c:\brem\archive" "c:\brem\log" "IN" 1
#powershell -ExecutionPolicy ByPass -File .\archiver.ps1 "c:\brem\log" "c:\brem\archive" "c:\brem\log" "LOG" 1

param (
    [string]$source = $(throw "-Source path is required."),
    [string]$target = $(throw "-Archive path is required."),
	[string]$logDir = $(throw "-Log path is required."),
	[string] $tag = $(throw "-File Tag is required."),
	[int] $days = $(throw "-Days old is required.")
)

$path = $source
$archPath = $target
$logPath = $logDir
$date = Get-Date -format yyyyMMddHHmm 

Write-Progress -activity "Archiving Data" -status "Progress:" 

$newArchPath = "$archPathDir\$date"
#Drop and create Archive directory
Remove-Item $newArchPath -Force -Recurse -ErrorAction SilentlyContinue
New-Item -ItemType directory -Path $newArchPath


If ( -not (Test-Path $newArchPath)) {ni $newArchPath -type directory} 
Get-Childitem -Path $path -recurse| Where-Object {$_.LastWriteTime -lt (get-date).AddDays(-$days)} | 
ForEach { 
	$fileName = $_.FullName
	try { 
		Move-Item $_.FullName -destination $newArchPath -force -ErrorAction:SilentlyContinue 
		"Successfully moved $fileName to $newArchPath" | add-content "$logPath\$date.log.txt" 
	} 
	catch { 
		"Error moving $_ " | add-content "$logPath\$date.log.txt" 
	} 
}


& "C:\Program Files\7-Zip\7z.exe" u -mx9 -t7z -m0=lzma2 ("$archPath\$tag.$date.7z") $newArchPath
if ($LASTEXITCODE -eq 0) {
	Remove-Item $newArchPath -Force -Recurse -ErrorAction SilentlyContinue
} else {
	Exit 99
}

#$files = dir $path -Recurse -Include $mask | where {($_.LastWriteTime -lt (Get-Date).AddDays(-$days).AddHours(-$hours).AddMinutes(-$mins)) -and ($_.psIsContainer -eq $false)}