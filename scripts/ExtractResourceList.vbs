'"cscript.exe" "c:\trem\bin/ExtractResourceList.vbs" "c:\trem\in/GDC Manila Resource List template v1 0.xlsx" "c:\trem\in/ResourceList.txt" "c:\trem\in/ManagersList.txt" "201604"
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
Const xlFilterCopy = 2 'Copy filtered data to new location.
Const xlFilterInPlace = 1 'Leave data in place.
Const xlCellTypeVisible = 12

'cscript c:\atri\bin\ExtractResourceList.vbs "C:\atri\in\GDC Manila Resource List template v1 0.xlsx" "ResourceList.txt" "ManagerList.txt" "201603"
Dim dtmStartTime: dtmStartTime = Timer

Wscript.echo "[" & Now & "]" & ": " & "==============================================================="
Wscript.echo "[" & Now & "]" & ": " & "Extracting Resource List"
Wscript.echo "[" & Now & "]" & ": " & "==============================================================="

Dim a_batchid: a_batchid = WScript.Arguments(3)
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

Dim a_resourceList: a_resourceList = WScript.Arguments(0)
'File checks
If CheckFileExist(a_resourceList) = False Then
	Wscript.Echo "[" & Now & "]" & ": " & a_resourceList & " not found!"
	WScript.Quit 99
End If

'Dim resourceList: resourceList = SpliceFileName(WScript.Arguments(0),C_FILENAME) & ".txt"
Dim resourceList: resourceList = WScript.Arguments(1)
DeleteFile resourceList
ExtractResourceList a_resourceList, resourceList
CleanResourceList resourceList, resourceList & ".tmp"
RenameFile resourceList & ".tmp", resourceList
DeleteFile resourceList & ".tmp"

'Dim managerList: managerList = SpliceFileName(WScript.Arguments(0),C_FILENAME) & ".Managers.txt"
Dim managerList: managerList = WScript.Arguments(2)
DeleteFile managerList
ExtractManagers a_resourceList, managerList
CleanManagers managerList, managerList & ".tmp"
RenameFile managerList & ".tmp", managerList
DeleteFile managerList & ".tmp"

Wscript.echo "[" & Now & "]" & ": " & "==============================================================="
Wscript.echo "[" & Now & "]" & ": " & "Script completed in: " & GetElapsedTime
Wscript.echo "[" & Now & "]" & ": " & "==============================================================="
Wscript.Quit

Function ExtractResourceList(inFile, outFile)
	Dim objExcel, objWorkbook, objWorksheet
	Dim objRange, objRange2
	
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = False
	objExcel.DisplayAlerts = False
	'Open Source data
	Set objWorkbook = objExcel.Workbooks.Open(inFile)
	Set objWorksheet = objWorkbook.Worksheets(1)
	
	'Sort it
	Set objRange = objWorksheet.UsedRange
	'Check if workbook has contents
	If objExcel.WorksheetFunction.CountA(objRange) = 0 Then 
        Wscript.echo "[" & Now & "]" & ": " & inFile & " is empty. Please check the file."
		objWorkbook.Close
		objExcel.Quit
		Set objWorksheet = nothing
		Set objRange = nothing
		Set objWorkbook = nothing
		Set objExcel = nothing
		Wscript.Quit 99
	End If 
	'Continue if ok
	Set objRange2 = objExcel.Range("D1")
	objRange.Sort objRange2, xlAscending, , , , , , xlYes
	'Save it
	objWorkbook.SaveAs outFile, xlText
	'Clean it
	objWorkbook.Close
	objExcel.Quit
End Function

Function ExtractManagers(inFile, outFile)
	'A: Employee QLID
	'B: Employee Name
	'C: Manager QLID
	'D: Manager Name
	'E: Onsite Flag

	Dim objExcel, objWorkbook, objWorksheet
	Dim objRange, objRange2, objRange3
	Dim rowCount
	
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = False
	objExcel.DisplayAlerts = False
	'Open Source data
	Set objWorkbook = objExcel.Workbooks.Open(inFile)
	Set objWorksheet = objWorkbook.Worksheets(1)
	'Drop unecessary columns. Succeeding column names will be relative to delete ones
	objWorksheet.Columns("A:B").EntireColumn.Delete
	objWorksheet.Columns("C").EntireColumn.Delete
	'Filter Uniquely
	rowCount = objWorksheet.Range("A1").CurrentRegion.Rows.Count
	Set objRange2 = objExcel.Range("A1:A" & rowCount)
	objRange2.AdvancedFilter xlFilterInPlace, , , True
	'Delete hidden rows
	Dim i
	For i = rowCount To 1 Step -1
		If objWorksheet.Rows(i).Hidden = True Then objWorksheet.Rows(i).EntireRow.Delete
	Next
	'Sort results
	Set objRange = objWorksheet.UsedRange
	objRange.Sort objRange2, xlAscending, , , , , , xlYes
	'Export to Text file
	objWorkbook.SaveAs outFile, xlText
	'Cleanup
	objWorkbook.Close
	objExcel.Quit
End Function

Function CleanResourceList(byval inFile, byval outFile)
    Dim inFso, outFso, reader, writer
	Dim inLine, outLine
	
    Set inFso = CreateObject("Scripting.FileSystemObject")
	Set outFso = CreateObject("Scripting.FileSystemObject")
	
	Dim i: i = 0
	If inFso.FileExists(inFile) Then
		Set reader = inFso.OpenTextFile(inFile, ForReading, True)
		Set writer = outFso.OpenTextFile(outFile, ForWriting, True)
		Do Until reader.AtEndOfStream
			'Strip Quotes and spaces
			outLine = ReplaceText(reader.Readline,"\s*" & chr(34),"|") 
			'Strip spaces
			outLine = ReplaceText(outLine,"\|\s*","|")
			'Add space to comma for names. Since Badge report
			'has a space after the comma
			outLine = ReplaceText(outLine,",",", ") 
			'Skip Header
			If i > 0  Then 
				writer.Writeline Trim(outLine)
			End If
			i = i + 1
		Loop
		reader.close
		writer.close
		Set inFso = Nothing
		Set outFso = Nothing
	Else
		WScript.Echo "[" & Now & "]" & ": " &  inFile & " Does not Exist"
	End If
End Function

Function CleanManagers(byval inFile, byval outFile)
	Dim inFso, outFso, reader, writer
	Dim inLine, outLine
	
    Set inFso = CreateObject("Scripting.FileSystemObject")
	Set outFso = CreateObject("Scripting.FileSystemObject")
	
	Dim i: i = 0
	If inFso.FileExists(inFile) Then
		Set reader = inFso.OpenTextFile(inFile, ForReading, True)
		Set writer = outFso.OpenTextFile(outFile, ForWriting, True)
		Do Until reader.AtEndOfStream
			'Strip Quotes and lead/trail spaces
			outLine = ReplaceText(reader.Readline,"\s*" & chr(34),"|") 
			outLine = Trim(ReplaceText(outLine,chr(34),""))
			outLine = ReplaceText(outLine,",",", ")
			'Skip Header
			If i > 0  Then 
				writer.Writeline Trim(outLine)
			End If
			i = i + 1
		Loop
		reader.close
		writer.close
		Set inFso = Nothing
		Set outFso = Nothing
	Else
		WScript.Echo "[" & Now & "]" & ": " &  inFile & " Does not Exist"
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

Function RenameFile(byval inFile, byval outFile)
	Dim oFSO 
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	If oFSO.FileExists(inFile) Then
        oFSO.CopyFile inFile, outFile, True
	Else
		WScript.Echo "[" & Now & "]" & ": " &  inFile & " Does not Exist"
	End If
	Set oFSO = Nothing
End Function

Sub DeleteFile(filespec)
   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")
   If  fso.FileExists(filespec) Then fso.DeleteFile(filespec)
End Sub

Function ReplaceText(byval inputStr, byval patrn, byval replStr)
	Dim objRegEx, str1
	' Create regular expression.
	Set objRegEx = New RegExp
	objRegEx.Global = True
	objRegEx.Pattern = patrn
	objRegEx.IgnoreCase = True
	' Make replacement.
	ReplaceText = objRegEx.Replace(inputStr, replStr)
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