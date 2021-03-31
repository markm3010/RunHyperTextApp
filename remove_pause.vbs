' File name: remove_pause.vbs
' Author: Mark Matthias
' Date: 08/12/2005

' Purpose:
  ' open \\some\directory\_download_TKSxx_Latest.bat (xx is 76 or 80)
 ' read line by line, if line does not say "Pause", output line to %tks_tests%\_download76_auto.bat


Dim FSO, TKS_TESTS, TKSDownloadScript, TKSDownloadScriptCaller
Dim ArgString
Dim I, ArgArray
Dim intArraySize
Dim FileHandleSrc, FileHandleDest, bMatch, objRegExp, fileNameOnly, nameAndPath
Dim DownloadFile, objRegExpVer, delFileFirst, RunVersionFile
Dim buildArg
CONST FOR_READING = 1
CONST OVERWRITE = True
CONST DOWNLOAD_TKS76_32 = "\\zeus\Applications\32Bit Optitest_download_TKS76_Latest.cmd.lnk"
CONST DOWNLOAD_TKS80_32 = "\\zeus\Applications\MMTest.cmd.lnk"
CONST DOWNLOAD_TKS80_64 = "\\zeus\Applications\64Bit Optitest_download_TKS80_Latest.cmd.lnk"
CONST RUNVER_FILE = "\runver.txt" ' duplicate decl in runtests.hta!
CONST DEST_FILE = "\_download_auto.bat"
CONST CALL_DEST_FILE = "\call_download_auto.bat"
Set WshShell = WScript.CreateObject("WScript.Shell")
TKS_TESTS = WshShell.ExpandEnvironmentStrings("%TKS_TESTS%")

TKSDownloadScript = TKS_TESTS & DEST_FILE
TKSDownloadScriptCaller = TKS_TESTS & CALL_DEST_FILE
RunVersionFile = TKS_TESTS & RUNVER_FILE

Set FSO = CreateObject("Scripting.FileSystemObject")
Set FileHandleVer = FSO.OpenTextFile(RunVersionFile)

Set objRegExpVer = New RegExp
objRegExpVer.Global = True
objRegExpVer.IgnoreCase = True
objRegExpVer.Pattern = "TKS76"

Line = FileHandleVer.ReadLine
bMatch = objRegExpVer.Test(Line)
	
If bMatch = True Then
	DownloadFile = DOWNLOAD_TKS76_32
Else
	DownloadFile = DOWNLOAD_TKS80_32
	objRegExpVer.Pattern = "_64" ' to catch TKS80_64 - there was not a TKS76_64 build...
	bMatch = objRegExpVer.Test(Line)
		
	if bMatch = True Then
		DownloadFile = DOWNLOAD_TKS80_64
	End If
End If
	
Line = FileHandleVer.ReadLine
objRegExpVer.Pattern = "^\s*build:\s*[0-9]{4,}(\.)?" ' This regex MUST match the one in runtests.hta !!! (search for "This regex MUST" to find the line)
	
bMatch = objRegExpVer.Test(Line)
If bMatch = True Then
	buildArg = Line
End If

FileHandleVer.Close
' ---------------------------------------------
' new stuff for 8.0
Set WshShell = WScript.CreateObject("WScript.Shell")
Set oShellLink = WshShell.CreateShortcut(DownloadFile)
Set FileHandleSrc = FSO.OpenTextFile(oShellLink.TargetPath, FOR_READING)
ArgString = oShellLink.Arguments
ArgArray = Split(ArgString , " ")
intArraySize = UBound(ArgArray, 1) ' one-dimensional array..

If (FSO.FileExists(TKSDownloadScriptCaller)) Then
	Set delFileFirst = FSO.GetFile(TKSDownloadScriptCaller)
	delFileFirst.Delete(True)
End If

Set FileHandleDest = FSO.CreateTextFile(TKSDownloadScriptCaller, OVERWRITE)

Line = TKSDownloadScript

For I = 0 To intArraySize
	Line = Line & " " & ArgArray(I)
Next

Line = Line & " " & buildArg
FileHandleDest.writeLine(Line)
' end new additions for 8.0
' ---------------------------------------------

If (FSO.FileExists(TKSDownloadScript)) Then
	Set delFileFirst = FSO.GetFile(TKSDownloadScript)
	delFileFirst.Delete(True)
End If

Set FileHandleDest = FSO.CreateTextFile(TKSDownloadScript, OVERWRITE)

Set objRegExp = New RegExp
objRegExp.Global = True
objRegExp.IgnoreCase = True
objRegExp.Pattern = "^\s*pause\s*$"
	
Do While FileHandleSrc.AtEndOfStream <> True
	Line = FileHandleSrc.ReadLine
	bMatch = objRegExp.Test(Line)
	If bMatch = False Then
		FileHandleDest.writeLine(Line)
	End If
Loop
FileHandleSrc.Close
FileHandleDest.Close
