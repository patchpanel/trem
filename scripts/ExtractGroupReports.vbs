'"cscript.exe" "c:\trem\bin/ExtractGroupReports.vbs" "c:\trem\in/ResourceList.txt" "c:\trem\in/ManagersList.txt" "c:\trem\in/201603 - BDG_TimeReport_V2.xlsx" "Summary" "Detailed Entry 
Option Explicit

Const C_FILENAME = 1
Const C_EXTNAME = 2
Const C_PATH = 3
' FileSystemObject.OpenTextFile
Const OpenAsDefault    = -2
Const CreateIfNotExist = -1
Const FailIfNotExist   = 0
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Const xlText = -4158
Const xlAscending = 1
Const xlDescending = 2
Const xlYes = 1
Const xlCellTypeVisible = 12

'args(0) - Extracted Resource List in text format
'args(1) - Manager List in text format
'args(2) - Badge Report in Excel Format
'args(3) - Badge Report sheet1 name - Summary
'args(4) - Badge Report sheet2 name - Detailed Entry Exit Pair
'args(5) - Badge Report sheet3 name - Detailed Raw
'args(6) - Report output Directory
'args(7) - Batch ID
'args(8) - Manager Report Tag
'+-------------------------------------------------------+'
'|                 MAIN ENTRY POINT                      |'
'+-------------------------------------------------------+'
Dim dtmStartTime: dtmStartTime = Timer
'Dim dtProcessed: dtProcessed = Left(IsoDate(Now),8)
Dim a_resourceList, a_managerList, a_badgeReport, a_srcSheet1, a_srcSheet2, a_srcSheet3, a_outputDir, a_batchid, a_MngrTag

a_resourceList = WScript.Arguments(0)
'a_managerList = SpliceFileName(a_resourceList, C_FILENAME) & ".Managers." & SpliceFileName(a_resourceList, C_EXTNAME)
a_managerList = WScript.Arguments(1)
a_badgeReport = WScript.Arguments(2)
a_srcSheet1 = WScript.Arguments(3)
a_srcSheet2 = WScript.Arguments(4)
a_srcSheet3 = WScript.Arguments(5)
a_outputDir = WScript.Arguments(6)
a_batchid = WScript.Arguments(7)
a_MngrTag = WScript.Arguments(8)

'Batch ID checking
Dim year, month, thedate
year = Left(a_batchid,4)
month = Right(a_batchid,2)
thedate =  "01-" & month & "-" & year
if IsDate(thedate) = false Then
	Wscript.Echo "Please enter a valid Batch ID (YYYYMM)"
	Wscript.Quit 99
End if
'Check for period close first
Dim appDataDir: appDataDir = GetAppDataDir
Dim grpLck, indLck
indLck = appDataDir & "\trem\tremind." & a_batchid & ".lck"
grpLck = appDataDir & "\trem\tremgrp." & a_batchid & ".lck"

If CheckFileExist(indLck) = True  And CheckFileExist(grpLck) = True Then
	Wscript.echo "[" & Now & "]" & ": Reporting Period " & thedate & " is already closed. Please Enter a new period to proceed."
	WScript.Quit 69
End If

Wscript.echo "[" & Now & "]" & ": " & "==============================================================="
Wscript.echo "[" & Now & "]" & ": " & "Extracting per Practice report"
Wscript.echo "[" & Now & "]" & ": " & "==============================================================="


If CheckFileExist(a_resourceList) = False Then
	Wscript.echo "[" & Now & "]" & ": " & a_resourceList & " not found!"
	WScript.Quit 99
End If

If CheckFileExist(a_managerList) = False Then
	Wscript.echo "[" & Now & "]" & ": " & a_managerList & " not found!"
	WScript.Quit 99
End If

If CheckFileExist(a_badgeReport) = False Then
	Wscript.echo "[" & Now & "]" & ": " & a_badgeReport & " not found!"
	WScript.Quit 99
End If

If CheckFolderExist(a_outputDir) = False Then
	Wscript.echo "[" & Now & "]" & ": " & a_outputDir & " does not exist!"
	WScript.Quit 99
End If

'Backup the Badge Report first
CopyFile a_badgeReport, SpliceFileName(a_badgeReport,C_FILENAME) & "_Backup" & a_batchid & "." & SpliceFileName(a_badgeReport,C_EXTNAME)

Wscript.echo "[" & Now & "]" & ": " & "================Mapping Employee to Manager================="
'Make the Excel and workbook object a global variable
'Workbook should be left open as processing occurs
'It might be faster than opening and closing it per Employee
Dim objExcel, objWorkbook
Dim recordsProcessed: recordsProcessed = 0
recordsProcessed = BuildPracticeReport(a_resourceList, a_badgeReport)
'Final Cleanup
Set objWorkbook = Nothing
Set objExcel = Nothing

Dim arrResults
arrResults = Split(recordsProcessed, "|", -1, 1)

Wscript.echo "[" & Now & "]" & ": " & "---------------------------------------------------------------"
Wscript.echo "[" & Now & "]" & ": " & "Employees with Managers: " & arrResults(0)
Wscript.echo "[" & Now & "]" & ": " & "Employees without Managers: " & arrResults(1)
Wscript.echo "[" & Now & "]" & ": " & "---------------------------------------------------------------"

Wscript.echo "[" & Now & "]" & ": " & "================Extraction of Practice Report=================="
recordsProcessed = ExtractPracticeReport(a_managerList, a_badgeReport, a_outputDir, a_srcSheet1, a_srcSheet2, a_srcSheet3)
arrResults = Split(recordsProcessed, "|", -1, 1)

Wscript.echo "[" & Now & "]" & ": " & "---------------------------------------------------------------"
Wscript.echo "[" & Now & "]" & ": " & "Managers with Employees: " & arrResults(0)
Wscript.echo "[" & Now & "]" & ": " & "Managers without employees: " & arrResults(1)
Wscript.echo "[" & Now & "]" & ": " & "---------------------------------------------------------------"

Wscript.echo "[" & Now & "]" & ": " & "==============================================================="
Wscript.echo "[" & Now & "]" & ": " & "Script completed in: " & GetElapsedTime
Wscript.echo "[" & Now & "]" & ": " & "==============================================================="
Wscript.Quit

Function BuildPracticeReport (byval resourceList, byval badgeReport)
	BuildPracticeReport = ""
	'Open the Excel app
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = False
	objExcel.DisplayAlerts = False
	Set objWorkbook = objExcel.Workbooks.Open(badgeReport)

	'Check if workbook has contents
	If (objExcel.WorksheetFunction.CountA("A1:K1") = 0 Or (CInt(objWorkbook.Worksheets.Count) <= 2)) Then 
        Wscript.echo "[" & Now & "]" & ": " & badgeReport & " is empty or has missing worksheets. Please check the file."
		objWorkbook.Close
		objExcel.Quit
		Set objWorkbook = nothing
		Set objExcel = nothing
		Wscript.Quit 99
	End If 
	'Continue if ok
	
	Dim oFso, reader
	'Parse the Resource List text file
	Set oFso = CreateObject("Scripting.FileSystemObject")
	If oFso.FileExists(resourceList) Then
		Dim arrLine, sLine, empName, mngrName
		Dim mCtr: mCtr = 0
		Dim nCtr: nCtr = 0
		Set reader = oFso.OpenTextFile(resourceList, ForReading, True)
		Do Until reader.AtEndOfStream
			sLine = reader.Readline
			arrLine = Split(sLine, "|", -1, 1)
			empName = Trim(arrLine(1))
			mngrName = Trim(arrLine(3))
			If MatchEmployee(empName, mngrName) = True then
				'WScript.echo "[" & Now & "]" & ": " & empName & Chr(9) & "Match"
				mCtr = mCtr + 1
			Else
				'WScript.echo "[" & Now & "]" & ": " & empName & Chr(9) & "*No Match*"
				nCtr = nCtr + 1
			End if
		Loop
		reader.close
		Set reader = nothing
		Set oFso = Nothing
	Else
		WScript.echo "[" & Now & "]" & ": " & resourceList & " Does not Exist"
	End If

	objWorkbook.Close True, badgeReport
	objExcel.Quit
	
	BuildPracticeReport = mCtr & "|" & nCtr
End Function

Function MatchEmployee(byval EmployeeName, byval ManagerName)
	Dim objWorksheet, srcRange, lastRow

	Dim fIsFound1: fIsFound1 = 0
	Dim fIsFound2: fIsFound2 = 0
	Dim fIsFound3: fIsFound3 = 0
	
	'==================================================
	'Summary Processing.
	'==================================================
	Set objWorksheet = objWorkbook.Worksheets(1)
	Set srcRange = objWorksheet.Range("A1:K1")
	'Add heading to Manager Column
	objWorksheet.Range("L1" & lastRow).value = "ManagerName"
	
	With objWorksheet
		If .AutoFilterMode = False Then srcRange.AutoFilter
		.Range("A1:K1").AutoFilter
		.Range("A1:K1").AutoFilter 3, "=" & EmployeeName
	End With
	fIsFound1 = objWorksheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count
	If fIsFound1 > 1 then
		lastRow = objWorksheet.Range("A1").CurrentRegion.Rows.Count
		objWorksheet.Range("L2:L" & lastRow).value = ManagerName
		If objWorksheet.AutoFilterMode Then objWorksheet.AutoFilter.ShowAllData
	End If
	
	'==================================================
	'Detailed Entry Exit Pair Processing.
	'==================================================
	Set objWorksheet = objWorkbook.Worksheets(2)
	Set srcRange = objWorksheet.Range("A1:K1")
	'Add heading to Manager Column
	objWorksheet.Range("L1" & lastRow).value = "ManagerName"
	
	With objWorksheet
		If .AutoFilterMode = False Then srcRange.AutoFilter
		.Range("A1:K1").AutoFilter
		.Range("A1:K1").AutoFilter 3, "=" & EmployeeName
	End With
	fIsFound2 = objWorksheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count
	If fIsFound2 > 1 then
		lastRow = objWorksheet.Range("A1").CurrentRegion.Rows.Count
		objWorksheet.Range("L2:L" & lastRow).value = ManagerName
		If objWorksheet.AutoFilterMode Then objWorksheet.AutoFilter.ShowAllData
	End If
	
	'==================================================
	'Detailed Raw Processing.
	'==================================================
	Set objWorksheet = objWorkbook.Worksheets(3) 
	Set srcRange = objWorksheet.Range("A1:H1")
	'Add heading to Manager Column
	objWorksheet.Range("I1" & lastRow).value = "ManagerName"
	
	With objWorksheet
		If .AutoFilterMode = False Then srcRange.AutoFilter
		.Range("A1:H1").AutoFilter
		.Range("A1:H1").AutoFilter 3, "=" & EmployeeName
	End With
	fIsFound3 = objWorksheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count
	If fIsFound3 > 1 then
		lastRow = objWorksheet.Range("I1").CurrentRegion.Rows.Count
		objWorksheet.Range("I2:I" & lastRow).value = ManagerName
		If objWorksheet.AutoFilterMode Then objWorksheet.AutoFilter.ShowAllData
	End If
	
	If fIsFound1 > 1 Or fIsFound2 > 1 Or fIsFound3 > 1 Then
		MatchEmployee = True
	Else
		MatchEmployee = False
	End If
	
	Set objWorksheet = Nothing
	Set srcRange = Nothing
End Function

Function ExtractPracticeReport(byval managerlist, _
							   byval srcExcelFile, _
							   byval outputDir, _
							   byval srcSheet1, _
							   byval srcSheet2, _
							   byval srcSheet3)
	Dim objExcel, objWorkbook
	Dim mngrName
	Dim oFso, reader
	
	ExtractPracticeReport = ""
	
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = False
	objExcel.DisplayAlerts = False
	Set objWorkbook = objExcel.Workbooks.Open(srcExcelFile)
	'Extract the per practice report
	Set oFso = CreateObject("Scripting.FileSystemObject")
	If oFso.FileExists(managerlist) Then
		Dim mCtr: mCtr = 0
		Dim nCtr: nCtr = 0
		Dim arrLine, sLine
		Set reader = oFso.OpenTextFile(managerlist, ForReading, True)
		Do Until reader.AtEndOfStream
			sLine = reader.Readline
			arrLine = Split(sLine, "|", -1, 1)
			mngrName = Trim(arrLine(1))
			If MatchManager(objExcel, _
						    objWorkbook, _
							outputDir, _
							mngrName, _
							srcSheet1, _
							srcSheet2, _
							srcSheet3) = True then
				'WScript.echo "[" & Now & "]" & ": " & mngrName & Chr(9) & "Match"
				mCtr = mCtr + 1
			Else
				WScript.echo "[" & Now & "]" & ": " & mngrName & Chr(9) & "*No Match*"
				nCtr = nCtr + 1
			End if
		Loop
		reader.close
		Set reader = nothing
		Set oFso = Nothing
	Else
		WScript.echo "[" & Now & "]" & ": " & managerlist & " Does not Exist"
	End If
	
	objWorkbook.Close
	objExcel.Quit
	Set objExcel = Nothing
	Set objWorkbook = Nothing
	
	ExtractPracticeReport = mCtr & "|" & nCtr
End Function

Function MatchManager(excelObject, _
					  workbookObject, _
					  byval outputDir, _
					  byval ManagerName, _
					  byval srcSheet1, _
					  byval srcSheet2, _
					  byval srcSheet3)
	Dim objWorksheet, srcRange
	Dim objTgtWorkbook, objTgtWorksheet
	
	Dim fIsFound1: fIsFound1 = 0
	Dim fIsFound2: fIsFound2 = 0
	Dim fIsFound3: fIsFound3 = 0
	
	'Let's create the output workbook
	Set objTgtWorkbook = excelObject.Workbooks.Add
	'==================================================
	'Summary Processing. Read and Copy to new Worksheet
	'==================================================
	Set objWorksheet = workbookObject.Worksheets(srcSheet1) 
	Set srcRange = objWorksheet.Range("A1:L1")
	With objWorksheet
		If .AutoFilterMode = False Then srcRange.AutoFilter
		.Range("A1:L1").AutoFilter
		.Range("A1:L1").AutoFilter 12, "=" & ManagerName
	End With
	
	fIsFound1 = objWorksheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count
	If fIsFound1 > 1 then
		Set objTgtWorksheet = objTgtWorkbook.Worksheets(1) 'Default Sheet1
		objTgtWorksheet.Name = srcSheet1
		objWorksheet.AutoFilter.Range.Copy objTgtWorksheet.Range("A1")
		objTgtWorksheet.Columns("L").EntireColumn.Delete
		objTgtWorksheet.Cells.EntireColumn.AutoFit
	End If
	
	'==================================================
	'Detailed Entry Exit Pair Processing.
	'==================================================
	Set objWorksheet = workbookObject.Worksheets(srcSheet2)
	Set srcRange = objWorksheet.Range("A1:L1")
	With objWorksheet
		If .AutoFilterMode = False Then srcRange.AutoFilter
		.Range("A1:L1").AutoFilter
		.Range("A1:L1").AutoFilter 12, "=" & ManagerName
	End With
	fIsFound2 = objWorksheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count
	If fIsFound2 > 1 then
		Set objTgtWorksheet = objTgtWorkbook.Worksheets(2) 'Default Sheet2
		objTgtWorksheet.Name = srcSheet2
		objWorksheet.AutoFilter.Range.Copy objTgtWorksheet.Range("A1")
		objTgtWorksheet.Columns("J:L").EntireColumn.Delete
		objTgtWorksheet.Cells.EntireColumn.AutoFit
	End If

	'==================================================
	'Detailed Raw Processing.
	'==================================================
	Set objWorksheet = workbookObject.Worksheets(srcSheet3) 
	Set srcRange = objWorksheet.Range("A1:I1")
	With objWorksheet
		If .AutoFilterMode = False Then srcRange.AutoFilter
		.Range("A1:I1").AutoFilter
		.Range("A1:I1").AutoFilter 9, "=" & ManagerName
	End With
	fIsFound3 = objWorksheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count
	If fIsFound3 > 1 then
		Set objTgtWorksheet = objTgtWorkbook.Worksheets(3) 'Default Sheet3
		objTgtWorksheet.Name = srcSheet3
		objWorksheet.AutoFilter.Range.Copy objTgtWorksheet.Range("A1")
		objTgtWorksheet.Columns("I").EntireColumn.Delete
		objTgtWorksheet.Cells.EntireColumn.AutoFit
	End If

	If fIsFound1 > 1 Or fIsFound2 > 1 Or fIsFound3 > 1 Then
		objTgtWorkbook.Close True, outputDir & "\" & a_batchid & Space(1) & a_MngrTag & Space(1) & ManagerName
		MatchManager = True
	Else
		MatchManager = False
	End If
	
	Set srcRange = Nothing
	Set objWorksheet = Nothing
	Set objTgtWorkbook = Nothing
	Set objTgtWorksheet = Nothing
End Function

Function RenameFile(byval inFile, byval outFile)
	Dim oFSO 
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	If oFSO.FileExists(inFile) Then
        oFSO.CopyFile inFile, outFile, True
	Else
		WScript.echo "[" & Now & "]" & ": " & inFile & " Does not Exist"
	End If
	Set oFSO = Nothing
End Function

Function CheckFileExist(byval path)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(path) Then
		Set fso = Nothing
		CheckFileExist = True
	Else
		Set fso = Nothing
		CheckFileExist = False
	End If
End Function

Function CheckFolderExist(byval fldr)
   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")
   If fso.FolderExists(fldr) Then
		Set fso = Nothing
		CheckFolderExist = True
   Else
		Set fso = Nothing
		CheckFolderExist = False
   End If
End Function

Function SpliceFileName(byval fname, byval mode)
If Len(fname) > 0 Then
	If mode = C_EXTNAME Then
		SpliceFileName = Right(fname,Len(Trim(fname)) - InStr(fname,"."))
	End If
	If mode = C_FILENAME Then
		SpliceFileName = Left(fname,Len(Trim(fname)) - (Len(Trim(fname)) - InStr(fname,".")+1))
	End If
	If mode = C_PATH Then
		SpliceFileName = Left(fname,InstrRev(fname,"\")-1)
	End If
End If
End Function

REM Function IsoDate(byval dt)
	REM IsoDate = ((year(dt)*100 + month(dt))*100 + day(dt))*10000 + hour(dt)*100 + minute(dt)
REM End Function

Function CopyFile(byval inFile, byval outFile)
	Dim oFSO 
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	If oFSO.FileExists(inFile) Then
        oFSO.CopyFile inFile, outFile, True
	Else
		WScript.echo "[" & Now & "]" & ": " & inFile & " Does not Exist"
	End If
	Set oFSO = Nothing
End Function

Function GetElapsedTime
    Const SECONDS_IN_DAY = 86400
    Const SECONDS_IN_HOUR = 3600
    Const SECONDS_IN_MINUTE = 60
    Const SECONDS_IN_WEEK = 604800
 
    Dim dtmEndTime: dtmEndTime = Timer
	Dim seconds, minutes, hours, days
	
    seconds = Round(dtmEndTime - dtmStartTime, 2)
    If seconds < SECONDS_IN_MINUTE Then
        GetElapsedTime = seconds & " seconds "
        Exit Function
    End If
    If seconds < SECONDS_IN_HOUR Then 
        minutes = seconds / SECONDS_IN_MINUTE
        seconds = seconds MOD SECONDS_IN_MINUTE
        GetElapsedTime = Int(minutes) & " minutes " & seconds & " seconds "
        Exit Function
    End If
    If seconds < SECONDS_IN_DAY Then
        hours = seconds / SECONDS_IN_HOUR
        minutes = (seconds MOD SECONDS_IN_HOUR) / SECONDS_IN_MINUTE
        seconds = (seconds MOD SECONDS_IN_HOUR) MOD SECONDS_IN_MINUTE
        GetElapsedTime = Int(hours) & " hours " & Int(minutes) & " minutes " & seconds & " seconds "
        Exit Function
    End If
    If seconds < SECONDS_IN_WEEK Then
        days = seconds / SECONDS_IN_DAY
        hours = (seconds MOD SECONDS_IN_DAY) / SECONDS_IN_HOUR
        minutes = ((seconds MOD SECONDS_IN_DAY) MOD SECONDS_IN_HOUR) / SECONDS_IN_MINUTE
        seconds = ((seconds MOD SECONDS_IN_DAY) MOD SECONDS_IN_HOUR) MOD SECONDS_IN_MINUTE
        GetElapsedTime = Int(days) & " days " & Int(hours) & " hours " & Int(minutes) & " minutes " & seconds & " seconds "
        Exit Function
    End If
End Function

Function GetAppDataDir
	Dim objShell, appDataDir
	Set objShell = CreateObject( "WScript.Shell" )
	appDataDir = objShell.ExpandEnvironmentStrings("%APPDATA%")
	Set objShell = nothing
	GetAppDataDir = appDataDir
End Function